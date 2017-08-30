# PWExp
Sending Emails to Users about PW that are about to expire

When you download, make sure that the HTML file is in the same folder as the PS1 file.
The EXE can reside anywhere as it is only used to review XML files.

# Next Steps and Testing
Start by saving the files to a folder.

Open PowerShell and change the location to the same folder you saved the files in.

Then run:
  & .\EmailNotifyWhenPwWillExpireSoon.ps1 -DoNotSendEmail -Verbose

This will let you test before sending emails to users.

Record any problems.

Let me know.


This should create a "Log Files" folder where the logs and XML information gets saved to.

Check there after the run to see the full log.
  
