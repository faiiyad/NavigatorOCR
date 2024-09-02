# Navigator OCR
OCR system that scans a Passport for specific data and updates a database automatically.

## Installing
**Step 1:** Ensure that your device has Python 3 set up properly (if not, click [here](https://www.python.org/downloads/release/python-3123/) to download Python 3).

**Step 2:** Download the Tesseract OCR program by clicking [here](https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe).

**Step 3:** Setup Tesseract OCR by adding it to the system's PATH ([tutorial here](https://www.youtube.com/watch?v=2kWvk4C1pMo&pp=ygUUdGVzc2VyYWN0IG9jciBweXRob24%3D)).

> [!TIP]
> You can check if tesseract works by typing `tesseract` in Command Promt.

**Step 4:** Run the following commands to download the external libraries needed for the program to work correctly:

```
pip install PySimpleGUI
pip install pytesseract
pip install opencv-python
pip install openpyxl
```
> [!NOTE]
> You need to set-up an account for PySimpleGUI to remove the watermark ([here](https://www.pysimplegui.com/)).

**Step 5:** Download the `ver1.py` file and save it to your computer.

**Step 6:** Run the `ver1.py` file.


## Common Problems and fixes
- With low resolution files, the OCR may have difficulty recognizing the MRZ text, so it is advised to crop the image to only include the MRZ part. Subsequently, only the image of only the MRZ part can be taken to save time.

- Due to the presence of Bangla characters, using OCR on the whole passport page will lead to errors in reading the text, so it is strongly advised to take a picture of the MRZ portion only.

>[!NOTE]
><img src="https://github.com/user-attachments/assets/6e5a6b59-13fd-4b4a-854d-6a8d3f50412e" width="200"><br/>
>The MRZ is always located at the bottom part of the passport page.


- Sometimes, the names are scanned incorrectly. While a text cleaning and parsing function has been included within the program, it is advised to always check the Surname and Given Name fields.




