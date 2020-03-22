# RISE_tools
Tool to convert slides from pptx format to jupyter notebook (ipynb) slides with RISE extension.

![Original slide in pptx format](readme_images/pptx.gif)

**Original slide in pptx format**

<br>

<br>

![Converted slides to jupyter notebook](readme_images/ipynb.gif)

**Converted slides to jupyter notebook**

# Usage

`python ppt2rise.py example_2/simple.pptx`

Creates the jupyter notebook `example_2/simple.ipynb` (with all images stored at `example_2/images`)

`python ppt2rise.py example_2/simple.pptx another_folder/another_name.ipynb`

Creates the jupyter notebook `another_folder/another_name.ipynb` (with all images stored at `another_folder/images`)
