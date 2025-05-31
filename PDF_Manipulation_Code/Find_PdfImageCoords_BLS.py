from PIL import ImageTk
from pdf2image import convert_from_path
import tkinter as tk

# Specify the path to the invoice you want to find the coords you'll put in the vouching_automation file.
pdf_path = r""

images = convert_from_path(pdf_path, poppler_path=r"external_items/poppler-24.08.0/Library/bin")


page_number = 0
image = images[page_number]
window = tk.Tk()
canvas = tk.Canvas(window)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)


scrollbar = tk.Scrollbar(window)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
canvas.configure(yscrollcommand=scrollbar.set)
scrollbar.configure(command=canvas.yview)
tk_image = ImageTk.PhotoImage(image)

canvas.create_image(0, 0, anchor=tk.NW, image=tk_image)


def handle_click(event):
    x = canvas.canvasx(event.x)
    y = canvas.canvasy(event.y)
    print(f"Clicked position - X: {x}, Y: {y}")




canvas.bind("<Button-1>", handle_click)
canvas.config(scrollregion=canvas.bbox(tk.ALL))

window.mainloop()
#Coords go from X:0 to 1700 and Y: 0 to 2200
