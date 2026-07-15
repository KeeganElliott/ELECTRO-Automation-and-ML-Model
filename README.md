\# ELECTRO Automation \& Assistive/Predictive Machine Learning Platform



An integrated automation platform for high-voltage bushing design, simulation, digital signal processing (DSP), feature extraction, and assistive/predictive machine learning using Integrated Engineering Software's \*\*ELECTRO\*\*.



The project automates much of the engineering workflow required to analyze high-voltage bushings, collect geometric and simulation data, generate machine-learning-ready datasets, and develop predictive models capable of evaluating future bushing designs.



\---



\## Overview



This repository combines:



\- ELECTRO simulation automation

\- Geometry extraction

\- DSP-based feature extraction

\- Dataset generation

\- Machine learning

\- Design recommendation support



The objective is to reduce manual engineering effort while producing standardized datasets that can be used to train predictive models capable of estimating simulation outcomes and assisting future bushing design optimization.



\---



\# Major Capabilities



\## Geometry Automation



\- Automatic extraction of bushing geometry

\- Conductor dimensions

\- Shield dimensions

\- Shell dimensions

\- SolidWorks feature extraction

\- Segment-based ELECTRO geometry extraction



\---



\## ELECTRO Automation



Automates many repetitive simulation tasks including



\- Material assignment

\- Boundary condition setup

\- Voltage assignment

\- Static analysis execution

\- Transient analysis setup

\- Result extraction

\- Simulation metadata collection



\---



\## PDF Information Extraction



Automatic extraction of engineering specifications from drawing packages.



Examples include



\- Creepage distance

\- Rated voltage

\- Basic Impulse Level (BIL)



\---



\## Digital Signal Processing



Extracts quantitative descriptors from simulated electric-field distributions.



Current techniques include



\- FFT analysis

\- Wavelet analysis

\- Peak detection

\- Spectral energy measurements

\- Frequency-domain descriptors

\- Spatial-domain statistics



\---



\## Dataset Generation



Automatically produces machine-learning-ready datasets.



Collected features include



\- Geometry

\- Material properties

\- Voltage information

\- Creepage distance

\- Electric field statistics

\- DSP features

\- Simulation metadata

\- Pass / Fail labels



\---



\## Machine Learning



The repository includes the framework for an assistive and predictive machine-learning system.



Current architecture



\- Elastic-Net Logistic Regression

\- Constrained Gradient Boosted Trees

\- Constrained Deep Forest

\- Regression surrogate models

\- Counterfactual recommendation engine



Future work includes



\- Bayesian optimization

\- Genetic optimization

\- Active learning

\- Automated design recommendation



\---



\# Repository Structure



```

Application/



├── automation\_application.py

├── electro\_automation.py

├── electro\_geometry.py

├── solidworks\_geometry.py

├── tier1.py

├── fft\_analysis.py

├── wavelet\_analysis.py

├── pdf.py

├── electro\_assistive\_predictive\_model.py

├── INSTALL.bat

├── RUN.bat

└── README.txt

```



\---



\# Installation



Clone the repository



```bash

git clone https://github.com/KeeganElliott/ELECTRO-Automation-and-ML-Model.git

```



Create a virtual environment



```bash

python -m venv .venv

```



Activate



Windows



```bash

.venv\\Scripts\\activate

```



Install requirements



```bash

pip install -r requirements.txt

```



\---



\# Workflow



Typical workflow



```

SolidWorks Geometry

&#x20;       │

&#x20;       ▼

Geometry Extraction

&#x20;       │

&#x20;       ▼

ELECTRO Automation

&#x20;       │

&#x20;       ▼

Simulation

&#x20;       │

&#x20;       ▼

Result Extraction

&#x20;       │

&#x20;       ▼

FFT / Wavelet Processing

&#x20;       │

&#x20;       ▼

Feature Engineering

&#x20;       │

&#x20;       ▼

Dataset Generation

&#x20;       │

&#x20;       ▼

Machine Learning

&#x20;       │

&#x20;       ▼

Predictive Design Assistance

```



\---



\# Project Goals



The long-term goal is to build an engineering assistant capable of



\- Predicting simulation outcomes

\- Explaining predicted failures

\- Recommending design improvements

\- Reducing the number of expensive simulations required during high-voltage bushing development



\---



\# Technologies



\- Python

\- Integrated Engineering Software ELECTRO

\- SolidWorks

\- NumPy

\- SciPy

\- Pandas

\- Matplotlib

\- PyWavelets

\- OpenPyXL

\- OpenCV

\- Scikit-Learn



\---



\# Current Development Status



Current focus



\- Data collection

\- Feature engineering

\- Model training

\- Validation

\- Assistive recommendation system



Future work



\- Physics-informed machine learning

\- Automated design optimization

\- Larger training datasets

\- Improved recommendation engine

\- Deep-learning experimentation



\---



\# Author



Keegan Elliott



Electrical Engineering



Southern Illinois University Edwardsville



Research interests



\- High-voltage insulation systems

\- Digital Signal Processing

\- Machine Learning

\- Computational Engineering

\- Engineering Automation

