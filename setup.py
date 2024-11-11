from setuptools import setup, find_packages

setup(
    name="doctranslator",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        'openai>=1.0.0',
        'python-docx>=0.8.11',
    ],
    author="Suyao Yong",
    author_email="your.email@example.com",
    description="A multi-language document translator using Grok API",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    url="https://github.com/suyaoyong/DocTranslator",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.7',
) 