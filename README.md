# History of Expenses
 
The purpose of this project is to automate the processing of family expenses by integrating expense statements extracted from three different banks into one main Excel file and applying categories to each transaction, so that they can be analysed and visualized using Microsoft Excel Pivots and Charts.<br>
This task is achieved through the integration of the bank statements with Microsoft Excel, using **openpyxl**.

## Technologies used

Code written in ***python***<br>
***openpyxl*** for Excel processing<br>
***xls2xlsx*** for xls to xlsx conversion<br>
***aspose.pdf*** for pdf to xslx conversion<br>
***Tkinter*** for user interaction<br>
***Gmail API*** for e-mail handling and statement downloading<br>

## Main functionalities

***Important notice!*** *All references to sensitive personal data have been changed in the code and in the following images and gif the sensitive information will be hashed out.*

### Frontend

1. By using tkinter, the user can first choose which bank to process:

    ![Tkinter Main Window](/assets/images/TkinterMainWindow.png "Tkinter Main Window")

2. The second window allows the user to *process their bank*, and by doing so, they will be prompted to select the statement for processing:

    ![Tkinter Secondary Window gif](/assets/images/SecondaryWindowAnimation.gif "Tkinter Secondary Window")

3. The statement will be integrated into one main Excel file which will be used as the basis for graphs and analysis:

    ![Main Excel File](/assets/images/ExcelOutput.png "Main Excel File")

4. In the end, the visualization and analysis will look like below:

    - Column chart type for overview:

        ![Column Chart Type](/assets/images/Chart1.png "Column Chart Type")

    - Timeline chart type:

        ![Timeline Chart Type](/assets/images/Chart2.png "Timeline Chart Type")

### Backend

In the background, there are a number of things that happen, but the main important ***achievements*** are the following:
- Gmail handling with the help of Gmail API: mailbox access, label management and statement downloading
- File conversion and integration of different types of statements into one main file
- Applying categories to each transaction corresponding to types of expenses or incomes (column E and F in the main Excel File, module *BankOperations* > *TransactionManagement.py*)
- Applying the Account Owner and processed bank (column G and H)
- Performing actions such as: formatting, removing duplicates, preparing amounts for integration and proper categorization (credit/debit), deleting unnecessary transactions, etc.
- Error handling for unforseen situations which might disturb the process, with proper display in the tkinter window:

    ![Info](/assets/images/Info1.png "Info")
    ![Error](/assets/images/Error.png "Error")

## Challenges & Looking at the Future

This project being my *first attempt to programming*, I encountered a number of challenges along the way and so there is definitely ***room for improvement***:
- The **classes can be designed better**, so that code duplication is avoided (e.g.: at the moment, error handling and file selection done separately for each bank)
- In some places global variables are defined. Need to find ways to **drop the global variables**. This improvement would most probably go hand in hand with the point above.
- tqdm used as a progress bar in the interpreter, but this cannot show in the GUI. Need to **implement a loading bar** showing the progress in the GUI as well.
- Using Microsoft Excel for graphs and visualization is not the cleanest idea, so looking forward to **implementing Matplotlib for visualization** (or another more suitable library).