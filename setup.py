from setuptools import setup

setup(name='xlhelper',
      version='0.1.0',
      description='Some small minimal helper functions for loading Excel',
      url='https://github.com/darth-veitcher/xlhelper',
      author='James Veitch',
      author_email='james@jamesveitch.com',
      license='MIT',
      packages=['xlhelper'],
      zip_safe=False,
      install_requires=open('requirements.txt', 'r').readlines(),
      python_requires='>=3, <4',
      )
