# Ressources
# 1. https://stackoverflow.com/questions/32908639/open-pil-image-from-byte-file

import sys
from pptx import Presentation
import json
from PIL import Image
import io

from IPython import embed

def dummy_cells():
    cells = [
                {
                "cell_type": "markdown",
                "metadata": {
                    "slideshow": {
                    "slide_type": "slide"
                    }
                            },
                "source": [
                    "This is a markdown and main"
                            ]
                },
                {
                "cell_type": "code",
                "execution_count": None,
                "metadata": {},
                "outputs": [],
                "source": [
                    "# This is code and \"None\""
                        ]
                },
                {
                "cell_type": "code",
                "execution_count": None,
                "metadata": {
                    "slideshow": {
                    "slide_type": "fragment"
                    }
                },
                "outputs": [],
                "source": [
                    "# Code again and fragment"
                ]
                },
                {
                "cell_type": "code",
                "execution_count": 1,
                "metadata": {
                    "slideshow": {
                    "slide_type": "slide"
                    }
                },
                "outputs": [],
                "source": [
                    "# code and \"Slide\""
                ]
                },
                {
                "cell_type": "markdown",
                "metadata": {
                    "slideshow": {
                    "slide_type": "fragment"
                    }
                },
                "source": [
                    "Markdown and fragment"
                ]
                },
                {
                "cell_type": "markdown",
                "metadata": {
                    "slideshow": {
                    "slide_type": "-"
                    }
                            },
                "source": [
                    "Markdown and \"-\"\n",
                    " \n",
                    "Latex: $\\alpha$\n"
                        ]
                }
            ]
    return cells

def create_notebook(cells):
    """
    Creates a jupyter notebook given a list of cells
    """
    notebook_dict = {}
    notebook_dict["cells"] = cells
    notebook_dict["metadata"] = {
                                "celltoolbar": "Slideshow",
                                "kernelspec": {
                                    "display_name": "Python 3",
                                    "language": "python",
                                    "name": "python3"
                                    },
                                "language_info": {
                                    "codemirror_mode": {"name": "ipython","version": 3},
                                    "file_extension": ".py",
                                    "mimetype": "text/x-python",
                                    "name": "python",
                                    "nbconvert_exporter": "python",
                                    "pygments_lexer": "ipython3",
                                    "version": "3.7.4"
                                    }
                                }
    notebook_dict["nbformat"] = 4
    notebook_dict["nbformat_minor"] = 2
    return notebook_dict

def save_image(filepath, blob):
    """
    """    
    image = Image.open(io.BytesIO(blob))
    image.save(filepath)
    return

def save(notebook_dict, ipynb_filepath):
    with open(ipynb_filepath, 'w') as fp:
        json.dump(notebook_dict, fp)
    return

def create_cell(cell_type="markdown", content="This is a markdown and main", slide_type="slide"):
    """
    """
    cell = {
                "cell_type": cell_type,
                "metadata": {
                    "slideshow": {
                    "slide_type": slide_type
                    }
                            },
                "source": content
            }  
    return cell    
    

def ppt2rise(input_filepath, output_filepath):
    # load a presentation
    prs = Presentation(input_filepath)

    # Initialize the cell list
    cells = []

    # Iterate through slides
    for i, slide in enumerate(prs.slides):

        for shape in slide.shapes:
            print("\n"," | ".join(dir(shape)))

        # Text cells
        for j, shape in enumerate([_ for _ in slide.shapes if "text" in dir(_)]):
            if j==0:
                slide_type="slide"
                content = "## " + shape.text
            else:
                slide_type="-"
                content = shape.text
            # Do something
            my_text = [_+"\n" for _ in content.split("\n")]
            my_cell = create_cell(cell_type="markdown", content=my_text, slide_type=slide_type)
            cells.append(my_cell)
        
        # Image cells
        for j, shape in enumerate([_ for _ in slide.shapes if "image" in dir(_)]):
            blob = shape.image.blob
            content_type = shape.image.content_type.split("/")[-1]
            image_name = "{}_{}.{}".format(i,j,content_type)
            save_image(image_name, blob)
            my_text = """|["{}"]({})""".format("Image",image_name)
            my_cell = create_cell(cell_type="markdown", content=my_text, slide_type="")
            cells.append(my_cell)

    print(cells)
    # Create notebook and save it
    ipynb = create_notebook(cells)
    save(ipynb, output_filepath)
    return

    

if __name__=="__main__":
    """
    cells = dummy_cells()
    ipynb = create_notebook(cells)
    save(ipynb, "test.ipynb")
    """
    if len(sys.argv)==3:
        ppt2rise(input_filepath=sys.argv[1], output_filepath=sys.argv[2])
    else:
        ppt2rise()


