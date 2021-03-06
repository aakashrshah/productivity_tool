==========================================
Welcome to CITSQ documentation!
==========================================

.. image:: https://travis-ci.org/sphinx-gallery/sphinx-gallery.svg?branch=master
    :target: https://travis-ci.org/sphinx-gallery/sphinx-gallery

.. image:: https://readthedocs.org/projects/sphinx-gallery/badge/?version=latest
    :target: https://sphinx-gallery.readthedocs.io/en/latest/?badge=latest
    :alt: Documentation Status

.. image::     https://ci.appveyor.com/api/projects/status/github/sphinx-gallery/sphinx-gallery?branch=master&svg=true
    :target: https://ci.appveyor.com/project/Titan-C/sphinx-gallery/history

``Sphinx-Gallery`` is a `Sphinx <http://sphinx-doc.org/>`_ extension that builds an HTML
gallery of examples from any set of Python scripts.

.. image:: _static/demo.png
   :width: 80%
   :alt: A demo of a gallery generated by Sphinx-Gallery

The code of the project is on Github: `Sphinx-Gallery <https://github.com/sphinx-gallery/sphinx-gallery>`_

Features of Sphinx-Gallery
==========================

* :ref:`create_simple_gallery` by automatically running Python files,
  capturing outputs + figures, and rendering
  them into rST files ready for Sphinx when you build the documentation.
  Learn how to :ref:`set_up_your_project`
* :ref:`embedding_rst`, allowing you to interweave narrative-like content
  with code that generates plots in your documentation. Sphinx-Gallery also
  automatically generates a Jupyter Notebook for each your example page.
* :ref:`references_to_examples`. Sphinx-Gallery can generate mini-galleries
  listing all examples that use a particular function/method/etc.
* :ref:`link_to_documentation`. Sphinx-Gallery can automatically add links to
  API documentation for functions/methods/classes that are used in your
  examples (for any Python module that uses intersphinx).
* :ref:`multiple_galleries_config` to create and embed galleries for several
  folders of examples.

.. _install_sg:

Installation
============

Install via ``pip``
-------------------

You can do a direct install via pip by using:

.. code-block:: bash

    $ pip install sphinx-gallery

Sphinx-Gallery will not manage its dependencies when installing, thus
you are required to install them manually. Our minimal dependencies
are:

* Sphinx
* Matplotlib
* Pillow

Sphinx-Gallery has also support for packages like:

* Seaborn
* Mayavi

Install as a developer
----------------------

You can get the latest development source from our `Github repository
<https://github.com/sphinx-gallery/sphinx-gallery>`_. You need
``setuptools`` installed in your system to install Sphinx-Gallery.

You will also need to install the dependencies listed above and `pytest`

To install everything do:

.. code-block:: bash

    $ git clone https://github.com/sphinx-gallery/sphinx-gallery
    $ cd sphinx-gallery
    $ pip install -r requirements.txt
    $ pip install -e .

In addition, you will need the following dependencies to build the
``sphinx-gallery`` documentation:

* Scipy
* Seaborn

Sphinx-Gallery Show: :ref:`examples-index`
------------------------------------------

.. toctree::
   :maxdepth: 2
   :caption: Using Sphinx Gallery

   getting_started
   syntax
   configuration

.. toctree::
   :maxdepth: 2
   :caption: Advanced usage and information

   advanced
   faq
   utils


.. toctree::
   :maxdepth: 2
   :caption: Example galleries

   auto_examples/index
   tutorials/index
   auto_mayavi_examples/index

.. toctree::
   :maxdepth: 2
   :caption: API and developer reference

   reference
   changes
   Fork sphinx-gallery on Github <https://github.com/sphinx-gallery/sphinx-gallery>


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
