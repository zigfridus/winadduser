WINADDUSER is a Java tool which helps to add new users in Microsoft Windows.

It gets users list from a .xls (Microsof Excel) file.
Also it can create home directory for each user.
It tested on Micosoft Windows 7 and Micosoft Windows Server 2012.
It uses Windows commands:
- net user /add;
- wmic;
- icacls.

At first you should fill in .xls file. If you want users home directories you should write path to the root home
folder in the B1 cell and the name of admins group in your language in the B2 cell.
If you don't want home directories write dash (-) in B1 and B2 cells.
Administrators group will get right for all users home directories.
After that you should write all information about new users. Don't leave any cell empty.
Possible values for columns are:
- active:{yes | no}
Activates or deactivates the account
- expires:{date | never}
Causes the account to expire if date is set. The never option sets no time limit on the account.
An expiration date is in the form mm/dd/yyyy or dd/mm/yyyy, depending on the country code.
- password change:{yes | no}
Specifies whether users can change their own password.
- password required:{yes | no}
Specifies whether a user account must have a password.
- password expires:{true | false}
Specifies whether a users password expires.

You should have Java installed. You can get here: https://java.com
Run program:
>java -jar winadduser-1.0-jar-with-dependencies.jar import.xls

You can use any name of .xls file as the last argument.
