from distutils.core import setup

setup(
    description="Chart Library for lpOD",
    license="GPLv3",
    name = "lpod_chart",
    packages=['chart'],
    package_dir={'chart': 'src'},
    package_data={'chart': ['templates/chart.otc']},
    )
