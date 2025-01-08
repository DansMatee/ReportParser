# Report Parser
A report parser built to streamline iso-tank depot reports into a single excel spreadsheet for easy manipulation.

## Languages
Functionality created with python. Batch scripts created for ease of use.

## Functions
- Parses all depot types, pdfs and xsxls.
- Adds each report as a new block within a single spreadsheet, with labeled gate in and gate out sections.
- Dates added for easy readability.
 
## Running locally
1. Clone the repository into a folder.
2. Create the following folder structure in the repo root.
```bash
├── reports
│   ├── ARC
│   ├── MED
│   ├── TSA
│   └── WTS
```
3. Create a virtual env with python ```python -m venv .venv```
4. Activate the virtual env ```.venv\Scripts\Activate```
5. Install the required packages ```pip install -r requirements.txt```
6. Once requirements are installed, drag and drop depot specific reports in each folder, and use run.bat, spreadsheet will be created in the root folder once finished.
7. To clear spreadsheet and all reports from reports folders, run clear.bat.

## Images
- Output spreadsheet with reports added. <br>
![image](https://github.com/user-attachments/assets/37936d4c-7565-4cf3-95ab-bc94cdc1a4dc)

