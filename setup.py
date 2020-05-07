from setuptools import find_packages, setup

setup(
    name="winsay",
    version="1.0",
    packages=find_packages(),
    license="Private",
    description="say in windows",
    author="sukhbinder",
    author_email="sukh2010@yahoo.com",
    entry_points={
        'console_scripts': ['say = winsay:main', ],
    },
    install_requires=["pywin32"],

)
