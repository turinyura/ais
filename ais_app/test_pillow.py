from PIL import Image

print("Pillow успешно импортирован!")
img = Image.new('RGB', (100, 100), color='red')
img.save('test.png')
print("Файл test.png создан.")