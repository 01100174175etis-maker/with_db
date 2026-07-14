"""
إعدادات التثبيت والتهيئة للبرنامج
نظام إدارة المبيعات والحسابات المتقدم
يدعم Windows و Linux و macOS
"""
from setuptools import setup, find_packages

setup(
    name="sales_app",
    version="1.0.0",
    description="نظام إدارة الحسابات والمبيعات المتطور مع قاعدة بيانات SQLite وواجهة PyQt5",
    author="Sales Team",
    author_email="sales@example.com",
    python_requires=">=3.8",
    packages=find_packages(),
    install_requires=[
        "PyQt5==5.15.9",
        "pandas==2.0.3",
        "openpyxl==3.1.2",
    ],
    entry_points={
        'console_scripts': [
            'sales_app=main:main',
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
        "License :: OSI Approved :: MIT License",
    ],
)
