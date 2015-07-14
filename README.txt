This program will not work in its current state.
The required inputs to the program are:
- an integer specifiying the minimum amount of days an employee can be past due
- two database tables
     - a list of people and their past dues
     - a list of people and their supervisors
These inputs were initially retrieved from a Microsoft Access form and database. An import feature was created to import semicolon delimited .txt files into Microsoft Access using xml based import specifications. See Command13_Click() and Command14_Click().
The main function of the program is run with Command_Click0(), which sets the required variables then executes the MassEmail function.
The output of this program is a group of emails sent to the supervisors of employees with past due training.
I created this program as a tool for my company. I decided to store it here as a template for others working on basic automated email systems.
The InitOutlook() and Cleanup() functions were taken from http://wgafa.blogspot.com/2012/05/programmatically-send-email-in-access.html with very slight modifications. The SendEmail() function was also taken from this site but with much more modification.
To avoid any complications I've removed all references to the company and the personal information of others.
This is my first undirected project, as such I had no style guides to stick to and hadn't developed a strong, consistent style of my own.