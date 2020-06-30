# Overwriting restriction set on Headless Chrome

Script shows on how to overwrite the existing permission by manually adding experimental option and desired capabilities function.

## Cause:

Chrome has security permission that doesn't allow any files to download in headless mode for which web console is always required to perform any file download. 

## Code:

By injecting the below codes we could overwrite the default restriction set at headless chrome in chromedriver.

`option.add_experimental_option("prefs", {
        "download.default_directory": "C:\\path\\to\\Downloads",
        "download.prompt_for_download": False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False,
        'safebrowsing.disable_download_protection': True,
    })`

`capabilities = DesiredCapabilities.CHROME.copy()
    capabilities['acceptSslCerts'] = True
    capabilities['acceptInsecureCerts'] = True`


## Modules/Packages/Library:

These are the modules used in the script -

* Selenium
* OS
* win32com.client
* pandas

`Selenium` module is core of this script since this controls whole webpage.

`OS` module is used to act as interface with underlying operating system depending on the OS in user's machine.

`win32com.client` used to provide access to outlook/emails which is used as medium to convey the status report.

`pandas` standard livrary modile built in python to provide access to create teh logs while executing the scripts.

# License

Copyright (C) 2020 Kailash Prabhu

This is made public just to show on how to overwrite built in restrictions/function in headless chrome/chromedriver. It's currently being used in our organization.

Happy to accept any pull and recode it based on your organization, the main reason why this is made public.