------------------------------------
webCal 3.55
------------------------------------
by Jason Abbott
e-mail: webcal@webott.com
url: http://webott.com/jason/webCal.html
readme updated: April 12, 2000

webCal 3.55 is an ASP based calendaring program distributed in a ZIP file that should include the following files:

webCal3_month.asp        Month view
webCal3_week.asp         Week view
webCal3_detail.asp       Detail on a single event
webCal3_edit.asp         Edit an event
webCal3_updated.asp      Updates database after event edit
webCal3_deleted.asp      Deletes event from database
webCal3_find.asp         Search form
webCal3_found.asp        Processes and displays query results
webCal3_print-month.asp  Generates month view for printing
webCal3_print-week.asp   Generates week view for printing
webCal3_mini.asp         Miniature calendar used for popup
webCal3_popup.js         JavaScript for popup calendar
webCal3_themes.inc       Included color themes for all pages
webCal3_showrecur.inc    Special formatting for different recurrence
webCal3_verify.inc       Checks to see if user has logged in
webCal3_login.asp        Login screen

webCal3_user-admin.asp   User administration form
webCal3_user-edit.asp    Edit user details
webCal3_user-updated.asp Updates database after user edit
webCal3_user-deleted.asp Deletes user from database

show_status.inc          Generates JavaScript to update status bar
data/webCal.mdb          Access 2000 database for storing events
data/webCal3_data.inc    Connects to database
images/*.gif             Calendar toolbar images

Each file contains individual documentation.

INSTALLATION
------------------------------------
Copy the files to a directory beneath the WWW root of your ASP compatible web server.  The name of the main webCal directory is unimportant but the names of the subdirectories /data and /images cannot be changed without also modifying the calendar scripts.  Also, the file names cannot be changed without modifying the scripts.

Once the files are copied you may create a link to either webCal3_month.asp or webCal3_week.asp, or both.  The other files are invoked internally or linked to from the main calendar pages.  Click on the "week" icon at the end of each week in the month view to switch to the week view.  To switch to the month view from the week view, click on the month name at the top.

GETTING STARTED
------------------------------------
Before you begin adding events you will need to add one or more user accounts.  To do so, click on the key icon at the top of the calendar and login as the administrator.  As shipped, the administrator's user name is "admin" with a password of "user".  It is strongly recommended that you change these values (see next section).

Once you have logged in as the administrator, a user management icon should replace the key icon in the main calendar view.  Click on this icon to enter the user management form.  To add a user, select "Add" and enter user details. 

Once you have added a user account you can select "Logout" from the main calendar view to logout of the administrator account and then select the key icon to log in as the new user.  Once you have logged in, click on any date to add an event to that date.  Alternatively, you can click on a date before logging in as the new user, and you will be prompted to login at that time.

CUSTOMIZING THE CALENDAR
------------------------------------
webCal allows you to easily change the date format and colors used throughout the entire calendar by editing one file, webCal3_themes.inc.  This file includes insructions and examples on how to adjust the date format and color themes.  

USER MANAGEMENT
------------------------------------
All the accounts except for the administrator's can be edited online.  To make changes to the administrator's account, you must edit the table "cal_users" in the webCal database.  This is meant as a security measure.

Other accounts can be edited by any user given "Administrator" (as opposed to "User") level access.  When deleting an account you have the option of erasing all events scheduled by that user or moving those events to another user.

SECURING YOUR DATA
------------------------------------
If you will be using webCal on a public server then some steps should be taken to secure the webCal database.  There are two ways to restrict access to your database: change permissions on the database so that unauthorized users cannot download it or move the database to a directory outside of your web root.

  OPTION ONE
If you want to leave the database in the default location, the /data folder under the main webCal files, then you may want to restrict access to your database by adjusting permissions.  You'll want to be careful not to confuse FILE SYSTEM permissions with WEB SERVER permissions.  The Internet account, typically IUSR_[MACHINENAME], must have read and write access to the database within the FILE SYSTEM in order for webCal to function.  This should be the default configuration, meaning you won't need to change it.

The permissions you DO want to change are those of the WEB SERVER.  This can be done through the Management Console.  Within the Management Console, right-click on the webCal database, select it's properties, and disable "read" access.  If anyone guesses the path to and name of your database, they won't be able to read (ie download) it.

  OPTION TWO
If you are able to move files to locations outside of the web root (often /inetpub/wwwroot) then you may want to move the webCal database to a directory that's not part of the web site.  This makes it impossible for anyone to guess the path to your database and download it since only folders beneath the web site are accessible to Internet users.  If you do move the database then you will need to update webCal3_data.inc so that it points to the new location.  For example, if you move the file to c:\mydata then you would need to change this line

  DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & Server.Mappath("data/" & DataName & ".mdb")

to

  DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & "c:/mydata/" & DataName & ".mdb"

------------------------------------
Thank you for purchasing webCal.  I welcome any questions or feedback that you may have.

Jason Abbott
webcal@webott.com