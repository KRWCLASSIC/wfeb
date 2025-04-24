from setuptools import setup

setup(
    name="wfeb",
    version="1.0",
    py_modules=["wfeb"],
    entry_points={
        'console_scripts': [
            'wfeb=wfeb:main',
        ],
    },
    author="KRWCLASSIC",
    author_email="classic.krw@gmail.com",
    description="A tool to remove edit protection from Word documents",
    long_description=open("readme.md").read(),
    long_description_content_type="text/markdown",
    keywords="word, document, protection, docx",
    url="https://github.com/KRWCLASSIC/wfeb",
    python_requires=">=3.6",
) 