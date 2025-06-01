from setuptools import setup, find_packages

setup(
    name="macpst-converter",
    version="1.0.0",
    description="macOS PST File Converter - Convert Outlook PST files to various formats",
    author="Ricardo Rodrigues",
    author_email="ricardo@ricardorodrigues.ca",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    python_requires=">=3.8",
    install_requires=[
        "reportlab>=4.0.0",
        "python-dateutil>=2.8.0",
        "cryptography>=3.0.0",
        "lxml>=4.6.0",
        "pillow>=8.0.0",
        "chardet>=4.0.0",
    ],
    extras_require={
        "full": [
            "libpff-python>=20240114",
        ],
    },
    entry_points={
        "console_scripts": [
            "macpst-converter=macpst.main:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: MacOS",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Communications :: Email",
        "Topic :: Utilities",
    ],
)