# Adobe PDFServices Python SDK

This project aims at extracting information from PDF invoices using Adobe PDF Services in python.

## Prerequisites
The application has the following requirements:
* Python : Version 3.6 or above, pip, pandas, etc.

## Authentication Setup

The api credentials file and corresponding private key file for the samples is ```pdfservices-api-credentials.json``` and ```private.key``` respectively. For running this project, add both the above files in the root directory of downloaded zip file at the end of creation of credentials via [Get Started](https://www.adobe.io/apis/documentcloud/dcsdk/gettingstarted.html?ref=getStartedWithServicesSdk) workflow.

## Installation

Install the dependencies for the project as listed in the ```requirements.txt``` file with this command:

    pip install -r requirements.txt

## Running the project
The code itself is in the ```src/extract_pdf.py```. A Hundred test files used by the project can be found in ```resources/```. When executed, it creates an ```output```
child folder under the project root directory to store their results in output/ExtractedData.csv file.

### Extract PDF

#### Structured Information Output Format
The output of final project is ExtractedData.csv file with all the data appended to the csv file in the required format by using pandas Dataframes.

NOTE: This code is running in a loop to parse 100 pdf files at once.
So each time it is deleting the created zip file and structuredData.json file and appending the data to the csv file.

##### Extract Text, Table Elements and append it to csv file.

Command to run the python file containing the code from the root directory under the project.

```$xslt
python src/extract_pdf.py
```
or

```$xslt
py src/extract_pdf.py
```

### Licensing

This project is licensed under the Apache2 License. See [LICENSE](LICENSE.md) for more information. 
