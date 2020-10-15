# AutoCredit   <img src="logo.ico" height="35">
Hi! AutoCredit is a  **python** application used for creating **credit lists** from Excel sheets containing sales data. The application scans the Excel sheet and extracts the rows where the full payment has not yet been received, all the extracted data along with the calculated total pending amount is compiled into a **PDF** for easier viewing and portability.  The generated PDF credit list could also be **mailed** to the concerned sales department within the application.
>**I created this application for a local textile distribution business for helping in improving their workflow.**  
>Note: The source code is tailor made to meet the specifications of the aforementioned business. Feel free to modify it to meet your specific needs.  
## Packages/Libraries used

- **tkinter** - GUI
- **openpyxl** - handling Excel files
- **win32com** - PDF conversion
## User Interface

The interface was created using **Tkinter** - Python's de-facto standard GUI (Graphical User Interface) package.  Yes, I know the interface looks like something from the 90s😅, but I didn't want to spend a lot of time on it and tkinter is a really old package (there are some limitations to what all you can achieve with it).  
<br>
<img src="/screenshots/ui-1.png" height="400"> <img src="/screenshots/ui-2.jpg" height="400"> 
<img src="/screenshots/ui-3.jpg" height="400"> <img src="/screenshots/ui-4.jpg" height="400">  
<img align ="centre" src="/screenshots/ui-5.jpg" height="400">  
<br>
> Now I'm planning on upgrading the GUI using the more modern **Kivy** python framework or **Electron JS**.  
## Exception handling

I have tried to make the application as straight forward and error free as possible. As the application might be used by people who might not have a technical background, a lot of work has been done to make the user experience real smooth.  
Graphical prompts stating necessary instructions have been implemented wherever the user might run into issues.  
- Such as, but not limited to :  
  - No network connection while emailing the credit list.
  - Excel is running in the background while generating credit list.
  - Input Excel file does not exist.
  - Credit list does not exist.  
  <br>
  <img src="/screenshots/exception-1.jpg" height="400"> <img src="/screenshots/exception-2.jpg" height="400">  
## Sample input/output
The program (in this case) has been written to handle Excel sheets of a specific business . Therefore, the format and order of the columns and values in the Excel files are important for the successful execution of the program.  

**Excel file used by the specific business :**  

  <img src="/screenshots/excel-1.jpg" height="400"><br><img src="/screenshots/excel-2.jpg" height="400">
  >Some rows have either empty or incomplete values under the "Payment Received" column (see the 2nd screenshot), these rows are extracted for generating the credit list.  
  
**Credit list generated for the specific business :**  

  <img src="/screenshots/pdf-1.jpg" height="770"><br><img src="/screenshots/pdf-2.jpg" height="600">  
  > The pending value for each row is calculated by computing **(Amount - Payment Received)**. The total pending amount is added at the bottom.  
## Thats all for now 🤍
This project has helped me a lot in developing my python & problem solving skills, and it was a lot of fun!  

Feel free to modify the source code to your hearts content.
