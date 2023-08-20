# Carbone.io-util

[Carbone](https://carbone.io) is a fantastic way to generate report starting from LibreOffice™ or Microsoft Office™ (ods, docx, odt, xslx...) documents.
All you have to do is to add a mustache-like placeholder `{d.companyName}` in your template.


# Table of content
<!-- TOC start (generated with https://github.com/derlin/bitdowntoc) -->

- [The problem](#the-problem)
- [The solution](#the-solution)
- [Minimum Requirements](#minimum-requirements)
  * [LibreOffice™](#libreoffice)
- [Getting started](#getting-started)
  * [LibreOffice™](#libreoffice-1)
- [Contributors](#contributors)

<!-- TOC end -->


<!-- TOC --><a name="the-problem"></a>
# The problem
You have a big template with alot of placeholder `{}`, i.e.:

![immagine](https://github.com/danibs/Carbone.io-util/assets/30932554/4d5c064e-8bbb-4904-a75b-869ff279d569)

It is very difficult to distinguish between text and placeholder. The more your template grows the more it seems a mess 😵‍💫.


<!-- TOC --><a name="the-solution"></a>
# The solution
The easiest thing to do is hightlight placeholders!

![immagine](https://github.com/danibs/Carbone.io-util/assets/30932554/4acd6873-2644-4ae1-b22a-69fcfdc1e028)


<!-- TOC --><a name="minimum-requirements"></a>
# Minimum Requirements
<!-- TOC --><a name="libreoffice"></a>
## LibreOffice™
Macro is tested on LibreOffice 7.5.x, but it probably works with an older version as well.


<!-- TOC --><a name="getting-started"></a>
# Getting started
<!-- TOC --><a name="libreoffice-1"></a>
## LibreOffice™
- download file `CarboneIO-LibreOffice_Module.bas`
- open LibreOffice™ Writer
- menu `Tools > Macros > Edit Macros...`

In the macros IDE:
- menu `File > Import BASIC...`
- select file `CarboneIO-LibreOffice_Module.bas`

Return to LibreOffice™ Writer:
- add some text and placeholders like:
  ```
  Dear {d.name} {d.surname},
  {d.years} years have passed since we met in {d.city}.
  ```
- menu `Tools > Macros > Run Macro...`
- choose `highlightCarbonePlaceholder`

  ![immagine](https://github.com/danibs/Carbone.io-util/assets/30932554/4691ff31-494e-4c0f-bce7-fc2674a1d3c9)


That is! 🥳


<!-- TOC --><a name="contributors"></a>
# Contributors
Many thanks to [KamilLanda](https://ask.libreoffice.org/u/kamillanda) 💪 that helped me so much with the LibreOffice™'s macro.
