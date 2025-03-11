# 库安装脚本
from setuptools import setup, find_packages

setup(
    name="leave_application_generator",
    version="1.0.0",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "pandas",
        "python-docx",
        "openpyxl",
    ],
    entry_points={
        "console_scripts": [
            "leaveapp=main:main",
        ],
    },
)
