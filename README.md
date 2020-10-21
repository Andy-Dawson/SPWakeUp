# SPWakeUp (SPWakeUp3)

SPWakeUp - Wake up your Sharepoint and WSS Sites

A simple command line application that touches each site collection and sub-site on a Sharepoint or WSS front-end server to rebuild the IIS cache.

---

**Note:** The original code is not mine, I'm providing this for those people still wanting to use SPWakeUp.

This project was originally hosted at Codeplex. The archive is hosted at https://archive.codeplex.com/?p=spwakeup, the original source code for each release is also avaialble within the archive file that can be downloaded from that location.

SPWakeUp was originally provided for SharePoint 2007. It was also made available for SharePoint 2010. I provided builds for SharePoint 2013 and 2016. I have used SPWakeUp numerous times and found it to be a simple and reliable solution to wake up a SharePoint on-premises farm before users start accessing it for the day.

If you wish to compile SPWakeUp yourself, please follow the information in [Build-it-yourself.md](Build-it-yourself.md).

---

Running SPWakeUp without any options will search for and find all Web Applications on your farm. Each Web App will in turn be searched for Site Collections and Sub-Sites. Once a list of every URL is built, spwakup will open an HTTP connection to each in turn. This rebuilds the caches on the server which improves performance for subsequent visits.

You'll generally want to schedule SPWakeUp to run once a day on your Web-Front End servers sometime after IIS has been reset. I usually just schedule mine for 4am. You can also run it manually any time you need to reset IIS.

Available run-time options are:

-Exclude: Excludes the listed Site Collection URL from being woken.
Can be used more than once. Example: SPWakeUp3.exe -Exclude:http://badsite.com -Exclude:http://badsite2.com

-Email: An email address that should be sent a log of the results.

-UserName: Name of the account that should be used to browse the sites. If no user name is set, sites are accessed under the current account. This option must be used in conjunction with the -Domain and -Password options.
-Domain: The domain of the account specified above.
-Password: The password that should be used to browse the sites.

-Authentication: The type of authentication used to browse the sites. Default value is NTLM.

-Verbose Shows the old-style verbose interface.