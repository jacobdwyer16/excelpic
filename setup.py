from setuptools import setup, find_packages


setup(
    name= "excelpic",
    version = "0.1.0",
    packages= find_packages(),
    author= "Jacob Dwyer",
    author_email="jacobdwyer16@gmail.com",
    description= "Convert Excel Ranges to Images with Ease",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
)