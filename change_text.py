search_text = "doc"
replace_text = "d"
  
with open(r'click.py', 'r') as file:
    data = file.read()
    data = data.replace(search_text, replace_text)
  
with open(r'main2.py', 'w') as file:
    file.write(data)

print("Funcionou")