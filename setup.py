from setuptools import setup

setup(name='pycourse',  # "pip install pycourse"
      version='0.1',    # unstable version
      description='Create NSC course narration docs with from pptx files',
      py_modules=['pycourse'],  # "import pycourse"
      url='http://github.com/storborg/funniest',
      author='Billy Buehl',
      author_email='billybuehl792@gmail.com',
      license='MIT',
      packages=['pycourse'],
      package_dir={'': 'src'},
      install_requires=[
          'markdown',
      ],
      zip_safe=False
)