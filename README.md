
# SRM 18th Regulation CO PO Attainment Calculator

The 18th Regulation CO PO Attainment Calculator Portal is a web-based application designed to streamline the administrative tasks of faculty members. This portal allows faculty to upload and manage PDF files, enter registration numbers, and generate CAM (Continuous Assessment Mark) tables for student evaluations. The user-friendly interface ensures efficient and organized handling of files and data, making it an essential tool for academic administration.

Features
Registration Number Input: Faculty can enter multiple registration numbers in a text area, facilitating the management of student data.

PDF File Upload: Allows for the upload of multiple PDF files, with the file names displayed dynamically as they are added.

CAM Table Generation: Automatically generates a CAM table based on the uploaded data, providing a structured format for assessing student performance.

Dynamic Interaction: Responsive buttons for generating CAM tables and calculating attainment, enhancing user interaction.

Downloadable Results: Option to download the generated results for offline access and record-keeping.


## Technologies Used
Frontend: HTML, CSS, JavaScript

Backend: Flask (Python)

Styling: Custom CSS for a polished and professional look

Deployment: Web-based application
## Usage/Examples

Enter Registration Numbers:In the provided text area, input one or more registration numbers separated by commas.

Upload PDF Files: 
 Use the file input to select and upload PDF files. Click the add button (+) to display the names of the uploaded files in a list.

Generate CAM Table:
 Click on the "Generate CAM" button to display the CAM table section.
The CAM table will be populated with placeholders for CO (Course Outcome) attainment.

Calculate Attainment:
 After generating the CAM table, click the "Calculate Attainment" button to process the data.
The "Download Results" button will then be displayed, allowing you to download the calculated results.


## Installation and Setup

To run tests, run the following command

1. Clone the repository:
```bash
git clone https://github.com/SRMVDP-SYNC/18_regulation-CO-PO-Automation.git
cd 18_regulation-CO-PO-Automation/18_regulation

```
2. Set up a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`

```
3. Install dependencies:
```bash
pip install -r requirements.txt

```
4. Run the project:
```bash
flask run

```
5. Access the portal:
Open a web browser and go to 'http://127.0.0.1:5000'


## Used By

This project is used by:

Faculty of Computer Science department, SRM Institute of Science and Technology, Vadapalani 


## Demo

This link contains the demonstration of our project:

https://youtu.be/npNrojmUdwY
