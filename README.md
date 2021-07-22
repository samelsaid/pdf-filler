# pdf-filler
A little program that I made for a friend with the help of some existing code. It can fill a pdf template from an excel sheet and also load images into the pdf.

## How to use it:
Some defaults are already set, you only need to pass a .xls file and pdf template to fill. Make sure the field names are the same in both.
If you want to add images, toggle the flag. You're going to need to pass the field name to lookup images by and the directory with the images in it.

Check "Needs work" for the current limitations and work-arounds

Make sure to install dependancies from requirements.txt --> ```pip install -r requirements.txt```

Either make it executable with ```chmod +x excel_to_pdf.py``` or run with ```python3 excel_to_pdf.py [args]```

-h and --help will have more details.

Enjoy!

## Needs work:
* Check for and auto create missing directories
  - either supply a directory to output to or create an excel_to_pdf directory. In either case, you're going to need a no_images directory inside of that if you're filling your template with images.
* Make useless files go away
  - if you choose to fill images, you need a no_images directory to store the text filled pdf. The images fill on top of these
* Images don't fill a "field"
  - Currently images fill in at the lowest layer of the pdf so any fields will end up covering the image


