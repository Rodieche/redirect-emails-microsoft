# Exchange Online Mail Rule GUI

A PowerShell script that provides a simple Windows Form graphical interface (GUI) to create scheduled mail flow rules (transport rules) in Exchange Online.

This tool is designed for administrators to easily configure mail redirection or copying for a specific mailbox with defined activation and optional expiry dates/times. It includes specific handling for the NZ Auckland timezone to ensure rules activate/expire precisely as intended.

## Features

* **User-Friendly GUI:** A simple Windows Form provides an intuitive interface for entering rule parameters.

* **Source & Destination:** Clearly specify the source mailbox (whose incoming mail is affected) and the target mailbox for redirection or copying.

* **Action Choice:** Select between two primary actions:

    * **Redirect:** Messages are sent ONLY to the target mailbox. The original recipient does NOT receive the message.

    * **Copy:** Messages are sent to BOTH the original recipient AND the target mailbox.

* **Scheduled Activation:** Set a precise date and time for the rule to become active. The script correctly converts the specified time in the NZ Auckland timezone (NZST/NZDT) to UTC for Exchange Online.

* **Optional Expiry:** Define an optional date and time for the rule to automatically expire and stop processing.

* **Exchange Online Integration:** Connects directly to Exchange Online PowerShell using your credentials to create the transport rule.

## Prerequisites

Before running this script, ensure you have the following:

* A Windows operating system with PowerShell installed.

* The **Exchange Online Management PowerShell module**. If you don't have it, open PowerShell as Administrator and run:

    ```
    Install-Module -Name ExchangeOnlineManagement -Force

    ```

* An account with sufficient permissions to connect to Exchange Online and create Transport Rules within your organization.

* You must run the PowerShell script with **Administrator privileges**.

## Installation

1.  Save the provided PowerShell script code as a `.ps1` file (e.g., `ConfigureMailRuleGUI.ps1`) on your local machine.

## How to Run

1.  Open **PowerShell as Administrator**. You can do this by searching for "PowerShell" in the Start menu, right-clicking, and selecting "Run as administrator".

2.  Navigate to the directory where you saved the `.ps1` file using the `cd` command. For example, if you saved it in your Documents folder:

    ```
    cd C:\Users\YourUsername\Documents

    ```

3.  Execute the script by typing the file name preceded by `.\`:

    ```
    .\ConfigureMailRuleGUI.ps1

    ```

4.  A Windows Form window titled "Configure Mail Forward/Copy Rule" will appear.

5.  Fill in the required fields:

    * **Source Mailbox:** The email address of the mailbox whose incoming messages the rule should apply to.

    * **Redirect/Copy To Mailbox:** The email address where the messages should be sent.

    * **Action:** Select either "Redirect" or "Copy".

    * **Activation Date and Time:** Choose the date and time (in NZ Auckland Time) when the rule should become active.

    * **Optional: Expiry Date and Time:** Check the box and select a date and time if you want the rule to automatically expire.

6.  Click the "Create Rule" button.

7.  If you are not already connected to Exchange Online in that PowerShell session, a Microsoft 365 sign-in window will appear. Enter your credentials for an account with the necessary permissions.

8.  The "Status" label at the bottom of the form will update to show the progress (Connecting, Creating Rule, Success, or Error).

## Troubleshooting

* **"Module not found" error:** Ensure you have installed the `ExchangeOnlineManagement` module using `Install-Module -Name ExchangeOnlineManagement -Force` in an elevated PowerShell session.

* **"Cannot connect to Exchange Online" / Authentication prompts:** Verify your internet connection and that you are using credentials for an account with permissions to connect to Exchange Online. Running PowerShell as Administrator is required.

* **"Parameter cannot be found" error:** This indicates an issue with the parameters being passed to `New-TransportRule`. Ensure you are using the latest version of the script and that the parameters match the current cmdlet documentation (though the provided script should be up-to-date).

* **"Error creating rule: Method invocation failed because..." (related to String methods):** This was a previous issue related to string formatting in the rule name. Ensure you are using the latest version of the script where this has been corrected.

* **Rule not appearing in Exchange Admin Center:** After successful creation (Status shows "Rule created successfully!"), it might take a few moments for the rule to appear in the Exchange Admin Center GUI. You can also verify using `Get-TransportRule` in PowerShell (see below).

* **Rule not working as expected:** Double-check the rule configuration in the Exchange Admin Center. Ensure the source and destination mailboxes are correct, the action is set as intended, and the activation/expiry dates are correct.

## Verifying the Rule

After the script reports success, you can verify the rule's creation using PowerShell:

1.  Keep the PowerShell session open or reconnect to Exchange Online.

2.  Run the following command, replacing `"source.mailbox@yourdomain.com"` with the actual source mailbox you entered:

    ```
    Get-TransportRule -Filter {SentTo -eq "source.mailbox@yourdomain.office365.com"} | Format-List Name, State, ActivationDate, ExpiryDate, SentTo, RedirectMessageTo, CopyTo

    ```

    This will list details of rules applying to that mailbox, allowing you to confirm your new rule exists and is configured correctly.
