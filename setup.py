from setuptools import setup

setup(
    name="main",
    version="0.1",
    py_modules=["main"],
    install_requires=[
        "Click",
    ],
    entry_points="""
        [console_scripts]
        save-guides-to-database=main:save_guides_to_database
        check-paid-guides=main:check_paid_guides
        find-paid-guides=main:find_paid_guides
    """,
)
