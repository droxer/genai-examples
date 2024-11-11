from llama_index.core import PromptTemplate
from llama_index.llms.openai import OpenAI
import re
import os
import json


prompt = """
    You are a power point presentation specialist. You are asked to create
    the content for a presentation about {topic}.
    You have been given the following information to create a presentation:
    ---
    {information}.
    ---
    Structure the information in a way that it can be put in a power point
    presentation. Each slide should have a title and content, with the content
    being a summary of the information provided. Each slide should have one or
    more sentences that capture the key points of the information.
    Return the structured information as a JSON as follows.
    Your answer should only contain the JSON - no markdown formatting.
 """

prompt_template = PromptTemplate(prompt)
prompt_examples = """
    Example:
    {"slides": [
    {"title": "Slide 1", "content": "Content for slide 1"},
    {"title": "Slide 2", "content": "Content for slide 2"},
    {"title": "Slide 3", "content": "Content for slide 3"},
    ]}
"""

test_topic = "the benefits of exercise"
test_information = """
Exercise plays a crucial role in maintaining both physical and mental health.
Engaging in regular physical activity can significantly reduce the risk of
chronic diseases such as heart disease, diabetes, and obesity. It also enhances
muscular strength, flexibility, and endurance. Beyond physical benefits, exercise
contributes to improved mental health by reducing symptoms of depression and anxiety,
boosting mood through the release of endorphins, and improving cognitive function.
It fosters a sense of well-being and can be a great way to manage stress.
Overall, incorporating exercise into one's daily routine is a key factor in
achieving a healthier and more balanced lifestyle.
"""

content_prompt = (
    prompt_template.format(topic=test_topic, information=test_information)
    + prompt_examples
)

api_key = os.getenv("OPENAI_API_KEY")
llm = OpenAI(model="gpt-4o", api_key=api_key)
slides = json.loads(llm.complete(content_prompt, True).text)


powerpoint_prompt = """
You are a PowerPoint presentation specialist. You'll get a list of slides, each
slide containing a title and content. You need to create a PowerPoint presentation
based on the provided slides.
But there is a catch: Instead of creating the presentation, provide python code
that generates the PowerPoint presentation based on the provided slides.
Use the package python-pptx to create the PowerPoint presentation.
The presentation should be visually appealing and professionally designed.

In your code, use a file called 'template.pptx' as the template for the presentation
and stick to the template's design.

If the slides content contains more than one information, make bullet points.
Save the presentation as 'presentation.pptx'.

Your answer should only contain the python code, no explanatory text.

Slides:

"""

presentation_code = llm.complete(powerpoint_prompt + str(slides), True).text

match = re.findall(r"python\n(.*?)\n```", presentation_code, re.DOTALL)
python_code = match[0]
exec(python_code)
