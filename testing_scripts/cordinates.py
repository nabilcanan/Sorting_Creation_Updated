from tkinter import Tk, filedialog, Canvas
from PIL import Image, ImageTk


def on_image_click(event):
    """Callback for mouse click on the image."""
    # Print the (x, y) coordinates of the clicked point
    print("Clicked at:", event.x, event.y)


def main():
    """Main function to open image and set up click interaction."""
    # Set up the main window
    root = Tk()
    root.title("Image Offset Finder")

    # Let user choose an image file
    file_path = filedialog.askopenfilename(title="Select an image",
                                           filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff")])

    if not file_path:
        print("No file selected.")
        return

    # Load the selected image using PIL
    img = Image.open(file_path)
    photo = ImageTk.PhotoImage(img)

    # Create a canvas to display the image
    canvas = Canvas(root, width=img.width, height=img.height)
    canvas.pack()

    # Add the image to the canvas
    canvas.create_image(0, 0, anchor="nw", image=photo)
    # Bind the mouse click event to the callback
    canvas.bind("<Button-1>", on_image_click)

    root.mainloop()


if __name__ == "__main__":
    main()
