
# Industrial Cost Accounting Process Modeling and Notation Using BPMN 2.0

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.24287293.svg)](https://doi.org/10.5281/zenodo.24287293)

**Zenodo Record:** [https://zenodo.org/records/24287293](https://zenodo.org/records/24287293)  
**DOI:** [10.5281/zenodo.24287293](https://doi.org/10.5281/zenodo.24287293)

---

## Overview

This project presents an educational and technical framework for modeling industrial cost accounting processes using **Business Process Model and Notation (BPMN) 2.0**.

It focuses on representing the relationships between manufacturing operations, data evaluation and analysis, and industrial cost accounting through structured process models.

The project combines process modeling artifacts with supporting visualization files and software utilities. Its purpose is to provide a structured reference for understanding, documenting, analyzing, and studying the flow of information and activities involved in industrial cost accounting.

The primary process model is provided in BPMN 2.0 format. Complementary Visio and HTML files are included to support visualization and review.

## Project Objectives

The main objectives are to:

- Model industrial cost accounting activities using BPMN 2.0.
- Represent the interaction between manufacturing, data evaluation and analysis, and cost accounting processes.
- Document the flow of operational and accounting information across organizational units.
- Represent reusable manufacturing operation cycles through BPMN Call Activities.
- Provide complementary visual representations of the process model.
- Include supporting utilities for inspecting Excel workbooks and converting BPMN and Visio process models.
- Support education, process documentation, and further development of industrial cost accounting models.

## Process Modeling Scope

The BPMN model covers three main process areas:

### 1. Manufacturing Process

Represents the manufacturing workflow, including the relevant organizational responsibilities, activities, events, and information exchanges.

The manufacturing process is organized into multiple lanes, representing the main operational and supporting functions involved in production.

### 2. Data Evaluation and Analysis Process

Represents the process area responsible for evaluating and analyzing operational data and providing relevant information for subsequent accounting activities.

### 3. Industrial Cost Accounting Process

Represents the industrial cost accounting process and its relationship with manufacturing operations and evaluated operational information.

### Reusable Manufacturing Operation Cycle

The model also includes a reusable manufacturing operation cycle designed to represent a structured sequence of operational activities.

Its main stages include:

- Preparation and setup
- First-piece production
- Inspection and evaluation
- Approval decision
- Batch production
- Recording and documentation
- Delivery
- Setup correction when required
- Return to the appropriate operational stage

The reusable cycle is referenced through BPMN Call Activities in the manufacturing process.

## Repository Structure

```text
Industrial Cost Accounting Process Modeling Education/
│
├── README.md
│
├── BPMN 2.0/
│   ├── Industrial_Cost_Accounting_Process_Modeling_and_Notation.bpmn
│   ├── Industrial_Cost_Accounting_Process_Modeling_and_Notation.vsdx
│   └── Industrial_Cost_Accounting_Process_Modeling_and_Notation.html
│
└── Analysis Scripts/
    ├── 01Excel Complete Content, Formula And Formatting Extractor [All Values And Format] [1405.08.25] [1].ps1
    ├── 02Excel Complete Content, Formula And Formatting Extractor [All Values [Comprehensive Report]] [1405.08.25] [1].ps1
    ├── 03Excel Complete Content, Formula And Formatting Extractor [All Values [Brief Report]] [1405.08.25] [1].ps1
    ├── 04BPMN 2.0 To Visio Converter [Initial Version] [1405.08.25] [1].py
    └── 05Visio To BPMN 2.0 Converter [Initial Version] [1405.08.25] [1].py
```

## Project Deliverables

### BPMN 2.0 Model

The `.bpmn` file contains the primary process model and is intended to serve as the main structured representation of the modeled workflows.

It includes BPMN elements such as:

- Collaboration and Participants
- Pools and Lanes
- Tasks and Activities
- Events and Gateways
- Sequence Flows
- Message Flows
- Call Activities
- Reusable Process Definitions

### Microsoft Visio Representation

The `.vsdx` file provides a visual representation of the process model in Microsoft Visio.

It is intended to make the process structure easier to review, discuss, and use in educational and documentation contexts.

### HTML Representation

The `.html` file provides an additional browser-viewable representation of the process model.

### Analysis and Conversion Scripts

The `Analysis Scripts` directory contains supporting PowerShell and Python utilities.

**Excel inspection utilities**

The PowerShell scripts are designed to extract and document different levels of Excel workbook information, including cell values, formulas, calculated results, and formatting-related information.

They provide supporting capabilities for workbook inspection and analysis. They are not intended to represent a complete industrial cost accounting or ERP system.

**BPMN and Visio conversion utilities**

The Python scripts provide initial conversion utilities for:

- BPMN 2.0 to Microsoft Visio
- Microsoft Visio to BPMN 2.0

These utilities are included to support experimentation, model interchange, and further development of process-model conversion workflows.

## Technical Notes and Limitations

- The BPMN 2.0 file is the primary structured process-model source.
- The Visio and HTML files are complementary representations.
- The reusable manufacturing operation cycle does not contain BPMN-DI layout coordinates in the supplied BPMN source. Therefore, its placement in a Visio representation may require a separately designed layout.
- The Python conversion utilities are initial versions. Their output should be reviewed and validated, particularly with respect to process elements, relationships, layout information, and model compatibility.
- Complete, lossless conversion between BPMN and Visio should not be assumed without dedicated validation.
- The PowerShell utilities are intended for workbook inspection and documentation. Their execution environment and dependencies may affect compatibility.
- This project is an educational and technical process-modeling framework. It is not a certified accounting system, a complete ERP solution, or a substitute for organization-specific accounting policies and controls.

## Language

The primary language of the process model and educational content is **Persian (Farsi)**.

Some technical descriptions, file names, and supporting documentation are provided in English.

## Intended Applications

This project may be useful for:

- Industrial cost accounting education
- Industrial engineering studies
- Business process modeling
- Manufacturing process documentation
- BPMN 2.0 learning and practical exercises
- Process analysis and documentation
- Academic and technical research
- Further development of industrial accounting workflows

## Citation and DOI

If you use, reference, or build upon this work, please cite the Zenodo record.

**Recommended citation:**

*Industrial Cost Accounting Process Modeling and Notation Using BPMN 2.0: A Persian-Language Educational Framework.* Zenodo.

**DOI:** https://doi.org/10.5281/zenodo.24287293

**Zenodo Record:** https://zenodo.org/records/24287293

The DOI provides a persistent identifier for referencing the archived project record.

## License and Usage

Please refer to the license specified in the Zenodo record for the applicable terms of use, redistribution, and modification.

Unless explicitly permitted by that license, the availability of project files in this repository should not be interpreted as unrestricted permission to reuse or redistribute them.

## Acknowledgment

This project brings together process modeling, industrial cost accounting concepts, manufacturing workflow documentation, and supporting software utilities in a unified educational framework.

Contributions, technical feedback, and suggestions for further development are welcome.

---

**Project:** Industrial Cost Accounting Process Modeling and Notation Using BPMN 2.0  
**Primary Standard:** BPMN 2.0  
**Primary Model Format:** `.bpmn`  
**Supporting Formats:** `.vsdx`, `.html`, `.py`, `.ps1`  
**Persistent Identifier:** [10.5281/zenodo.24287293](https://doi.org/10.5281/zenodo.24287293)