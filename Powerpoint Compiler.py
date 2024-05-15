######################################################################
# Author: Sam McFarland
#
# Purpose: A program which takes photos and compiles them in different folders for use with different donors.
######################################################################

from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import os


def create_desktop_folder(folder_name):
    # Get the path to the user's home directory
    path_to_user = os.path.expanduser("~")

    # Combine with the Desktop folder
    path_to_desktop = os.path.join(path_to_user, "Desktop")

    # Create the folder
    new_folder_path = os.path.join(path_to_desktop, folder_name)
    os.makedirs(new_folder_path)


def resize_image(image):
    """
    Resizes an image
    :param image: The path of the image to resize
    :return: None
    """

    # Open an image
    im = Image.open(image)

    # Specify the new size
    new_size = (1152, 648)

    # Resize the image
    resized_image = im.resize(new_size)

    # Save the resized image
    resized_image.save(image)


def construct_pp():
    """
    Constructs the not-personalized base powerpoint using photos that should be included in each powerpoint.
    :return: None
    """
    # Giving Image path for each image
    img_path = 'Photos/Construction Photo 1.jpeg'
    img2_path = 'Photos/Construction Photo 2.jpg'
    img3_path = 'Photos/Construction Photo 3.jpg'
    img4_path = 'Photos/Construction Photo 4.jpg'

    # Creating a Presentation object
    ppt = Presentation()

    # Sets the dimensions for the Presentation
    ppt.slide_width = Inches(16)
    ppt.slide_height = Inches(9)

    # Selecting blank slide
    blank_slide_layout = ppt.slide_layouts[6]

    # Attaching slides to ppt
    slide = ppt.slides.add_slide(blank_slide_layout)
    slide2 = ppt.slides.add_slide(blank_slide_layout)
    slide3 = ppt.slides.add_slide(blank_slide_layout)
    slide4 = ppt.slides.add_slide(blank_slide_layout)

    # For no margins
    left = top = Inches(0)

    # adding images to each slide
    slide.shapes.add_picture(img_path, left, top)
    slide2.shapes.add_picture(img2_path, left, top)
    slide3.shapes.add_picture(img3_path, left, top)
    slide4.shapes.add_picture(img4_path, left, top)

    # save file
    ppt.save('base.pptx')


def add_donor(donor):
    """
    A function for adding specific donor photos to the base powerpoint, saving them in a specific file.
    :param donor: The name of the donor photo to be added.
    :return:
    """
    # Sets the margins to 0
    left = top = Inches(0)

    # Sets the image path based on the donor parameter
    img_path = 'Personal_photos/' + donor + '.jpg'

    # Sets the presentation to the base created in construct_pp()
    ppt = Presentation("base.pptx")

    # Sets a slide layout and adds the donor photo to a new slide
    blank_slide_layout = ppt.slide_layouts[6]
    slide = ppt.slides.add_slide(blank_slide_layout)
    slide.shapes.add_picture(img_path, left, top)

    # Sets the path to save the new powerpoint and saves it
    path = "C:/Users/mcfarlands/Desktop/Donor Power Points/"
    ppt.save(path + donor + ".pptx")


def main():
    """
    A function that resizes the images to the correct size and creates powerpoints from them.
    :return: None
    """
    create_desktop_folder("Donor Power Points")
    # Resizes the images to the correct size
    resize_image("Photos/Construction Photo 1.jpeg")
    resize_image("Photos/Construction Photo 2.jpg")
    resize_image("Photos/Construction Photo 3.jpg")
    resize_image("Photos/Construction Photo 4.jpg")
    resize_image("Personal_photos/Jack.jpg")
    resize_image("Personal_photos/Joan.jpg")
    resize_image("Personal_photos/Koning.jpg")

    # Constructs the base powerpoint and a powerpoint for each donor
    construct_pp()
    add_donor("Jack")
    add_donor("Joan")
    add_donor("Koning")

    print("Your powerpoints have been created!")


if __name__ == "__main__":
    main()
