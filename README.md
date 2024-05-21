# Employee Time Card App and Shift Calculator
## Overview
This repository contains an Employee Time Card App and Shift Calculator designed to efficiently manage employee shifts and calculate total hours worked. The application is divided into two main components:

timecard.py: The main application used by employees to log their working hours.
admin.py: The administration module for adding, removing, and editing staff details, and calculating total hours worked.
Additionally, the repository includes a batch file to install dependencies and an HTML file for usage instructions.

## Features
### Timecard Module (timecard.py)
Clock In/Clock Out: Employees can log their working hours by clicking their respective buttons.
Visual Feedback: The application provides visual feedback and confirmation messages upon clocking in and out.
Data Storage: Working hours are stored in an Excel file, organized by month and employee.
### Admin Module (admin.py)
Staff Management: Administrators can add, remove, and edit staff details.
Hour Calculation: Automatically calculates total hours worked for each employee and updates the Excel sheet.
Menu Interface: Simple text-based menu for navigating the administrative functions.

## Installation
Clone the repository:
bash```
git clone https://github.com/yourusername/employee-timecard-app.git
```
Navigate to the project directory:
bash```
cd employee-timecard-app
```
Run the batch file to install dependencies:
bash```
./install_dependencies.bat
```
## Usage
Running the Application
Timecard Module:
bash```
python timecard.py
```

Admin Module:
bash```
python admin.py
```

## Detailed Instructions
Refer to the readme.html file in the project directory for detailed instructions on how to use the application, including how to:

Add, remove, and edit staff in the admin module.
Log hours in the timecard module.
Calculate total hours worked for each employee.

## File Structure
bash```
employee-timecard-app/
│
├── timecard.py            # Main application file for employees
├── admin.py               # Admin module for managing staff and calculating hours
├── install_dependencies.bat  # Batch file for installing dependencies
├── readme.html            # Detailed usage instructions
├── README.md              # This file
└── requirements.txt       # List of dependencies (if needed)
```
## Contributing
Fork the repository.
Create your feature branch (git checkout -b feature/your-feature).
Commit your changes (git commit -am 'Add your feature').
Push to the branch (git push origin feature/your-feature).
Create a new Pull Request.
## License
This project is licensed under the MIT License. See the LICENSE file for details.


