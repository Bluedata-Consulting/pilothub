from setuptools import setup, find_packages

setup(
    name='pilothub',
    version='0.1.8',
    packages=find_packages(),
    install_requires=[
        # List your project dependencies here.
        # For example: 'numpy>=1.19.2'
        'openai==1.14.1',
        'python-pptx==0.6.23',
        'python-docx==1.1.0',
        'comtypes==1.2.0',

    ],
    # Add more information about your package
    author='Blue Data Consulting',
    author_email='anshu@bluedataconsulting.com',
    description='A short description of your project',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/Bluedata-Consulting/pilothub',
)
