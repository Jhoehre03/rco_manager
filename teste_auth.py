from rco.auth import conectar_chrome

browser = conectar_chrome()
print("Conectado! Título da página:", browser.title)