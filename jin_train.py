from jinja2 import Template

name = 'Lindy Booth'

dct = [{'name': 'Lindy Booth', 'age': 41}, {'name': 'Lindsey Striling', 'age': 36}, {'name': 'Eva Green', 'age': 50}]

text =""" {{% for girl in girls%}}
"""

tm = Template("Во славу {{ name }}")

print(tm.render(name=name))
