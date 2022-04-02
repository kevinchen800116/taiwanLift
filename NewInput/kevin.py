from PIL import ImageGrab

# 類似windows的printscreen
ImageGrab.grab().show()
clip=ImageGrab.grabclipboard()

clip.save("python.png")