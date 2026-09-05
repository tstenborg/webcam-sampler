# Meteor candidate observations from automated weather camera sampling in VBA

[![super-linter](../../actions/workflows/super-linter.yml/badge.svg)](../../actions/workflows/super-linter.yml) ![human-only code](https://img.shields.io/badge/human--only-code-white)

Automated sampling of webcam imaging was investigated as a means of meteor data
collection. A sampling system was developed in Excel with Visual Basic for
Applications (VBA), the scripting language built into Microsoft's Office suite.

---

<figure>
  <img src="assets/webcam-collage.png" alt="Sampling webcams managed by Airservices Australia on Norfolk Island." width="416">
  <figcaption>Figure 1. Sampling webcams managed by Airservices Australia on Norfolk Island.</figcaption>
</figure>

---

## Table of Contents

- [Key Files](#key-files)
- [Software Requirements](#software-requirements)
- [Quality Assurance](#quality-assurance)
- [Getting Started](#getting-started)
- [Acknowledgements](#acknowledgements)
- [References](#references)

## Key Files

| File                                | Notes                         |
| :---------------------------------- | :---------------------------- |
| `src/automated-webcam-sampler.xlsm` | Macro-enabled Excel workbook. |
| `src/webcam-sampler.bas`            | VBA module.                   |

The Excel workbook runs a standalone automated webcam sampler. Its constituent
VBA module, `webcam-sampler.bas`, has also been provided as separate file for
a) easy code review outside Excel and b) importing into another Excel workbook,
if desired.

## Software Requirements

- Excel (Windows version).

## Quality Assurance

The webcam sampler has been tested with Excel, version 2508, on Windows 11.

## Getting Started

### Sampler Configuration

The webcam sampler has default settings for target webcams, sampling cadence,
sampling period, download directory, etc. These should be configured in VBA to
meet your needs prior to executing the sampler. A modest, not advanced, level
of VBA knowledge is needed for this configuration.

### Execution Options

Ensure VBA macro execution is enabled in Excel.

Additionally, ensure Windows isn't also blocking the workbook's macros. To
resolve this, right click on the file to open the file's Properties. Go to the
General tab, Security section and select the Unblock checkbox.

Once complete, there are three main ways to execute the sampler.

#### 1. Macro Button

The single Excel worksheet features a "Take Shots" button at the top left of
the worksheet that will execute the sampler.

#### 2. VBA Editor

The DownloadFileAPI function of the VBA code will execute webcam sampling. This
can be triggered from Excel's VBA editor. The VBA editor can be opened by
selecting Developer, Visual Basic from the Excel ribbon. If the Developer
option isn't visible in the Excel ribbon, it can be added by selecting File,
Options, Customize Ribbon from the Excel main menu. (Excel version 2508 menu
options.)

#### 3. Imported Module

To execute the webcam sampler from another Excel workbook, a VBA module,
`webcam-sampler.bas`, is available for import and execution. The sampler
configuration options will still need to be customised to meet your sampling
needs.

## Acknowledgements

This work used publicly-accessible infrastructure managed by the Australian
government's Bureau of Meteorology and Airservices Australia.

## References

1. T. N. Stenborg, "Meteor candidate observations from automated weather camera
   sampling in VBA", in _Proc. International Meteor Conf._, U. Pajer,
   J. Rendtel, M. Gyssens and C. Verbeeck, Eds., Bollmannsruh, Germany, Oct.
   3&ndash;6, 2019, pp. 189&ndash;190.\
   [View PDF](https://articles.adsabs.harvard.edu/pdf/2020pimo.conf..189S.pdf)
   &nbsp; [SciX](https://scixplorer.org/abs/2020pimo.conf..189S/abstract)
