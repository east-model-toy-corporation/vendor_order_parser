from PIL import Image

# Convert ChatGPTImage_cat.png to icon.ico (multiple sizes)
src = "ChatGPTImage_cat.png"
dst = "icon.ico"

sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
img = Image.open(src)
img.save(dst, format="ICO", sizes=sizes)
print(f"Wrote {dst}")
