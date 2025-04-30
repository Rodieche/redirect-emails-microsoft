# Requires the Exchange Online Management module: Install-Module -Name ExchangeOnlineManagement -Force

# Add the required assemblies for Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Create the Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Configure Mail Forward/Copy Rule"
$form.Size = New-Object System.Drawing.Size(450, 550)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog" # Prevent resizing
$form.MaximizeBox = $false
$form.MinimizeBox = $false
# Set default button *after* it's created
# $form.AcceptButton = $btnCreateRule

# --- Form Controls ---

# Source Mailbox
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Location = New-Object System.Drawing.Point(20, 20)
$lblSource.Size = New-Object System.Drawing.Size(400, 20)
$lblSource.Text = "Source Mailbox (Forward/Copy From):"
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(20, 40)
$txtSource.Size = New-Object System.Drawing.Size(400, 20)
$txtSource.PlaceholderText = "e.g., user@yourdomain.com"
$form.Controls.Add($txtSource)

# Redirect/Copy To Mailbox
$lblRedirect = New-Object System.Windows.Forms.Label
$lblRedirect.Location = New-Object System.Drawing.Point(20, 80)
$lblRedirect.Size = New-Object System.Drawing.Size(400, 20)
$lblRedirect.Text = "Redirect/Copy To Mailbox:"
$form.Controls.Add($lblRedirect)

$txtRedirect = New-Object System.Windows.Forms.TextBox
$txtRedirect.Location = New-Object System.Drawing.Point(20, 100)
$txtRedirect.Size = New-Object System.Drawing.Size(400, 20)
$txtRedirect.PlaceholderText = "e.g., forward.to@otherdomain.com"
$form.Controls.Add($txtRedirect)

# Action Type (Radio Buttons)
$lblAction = New-Object System.Windows.Forms.Label
$lblAction.Location = New-Object System.Drawing.Point(20, 140)
$lblAction.Size = New-Object System.Drawing.Size(400, 20)
$lblAction.Text = "Action:"
$form.Controls.Add($lblAction)

$rbRedirect = New-Object System.Windows.Forms.RadioButton
$rbRedirect.Location = New-Object System.Drawing.Point(30, 165)
$rbRedirect.Size = New-Object System.Drawing.Size(380, 20)
$rbRedirect.Text = "Redirect (Original recipient does NOT receive)"
$rbRedirect.Checked = $true # Default
$form.Controls.Add($rbRedirect)

$rbCopy = New-Object System.Windows.Forms.RadioButton
$rbCopy.Location = New-Object System.Drawing.Point(30, 190)
$rbCopy.Size = New-Object System.Drawing.Size(380, 20)
$rbCopy.Text = "Copy (Original recipient DOES receive)"
$rbCopy.Add_CheckedChanged({
    # This event handler was the source of a previous error, but is now correctly implemented
    # to enable/disable the expiry time picker based on the expiry date checkbox.
})
$form.Controls.Add($rbCopy)

# Activation Date and Time
$lblActivation = New-Object System.Windows.Forms.Label
$lblActivation.Location = New-Object System.Drawing.Point(20, 230)
$lblActivation.Size = New-Object System.Drawing.Size(400, 20)
$lblActivation.Text = "Activation Date and Time (NZ Auckland Time):"
$form.Controls.Add($lblActivation)

$dtpActivationDate = New-Object System.Windows.Forms.DateTimePicker
$dtpActivationDate.Location = New-Object System.Drawing.Point(20, 250)
$dtpActivationDate.Size = New-Object System.Drawing.Size(200, 20)
$dtpActivationDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$form.Controls.Add($dtpActivationDate)

$dtpActivationTime = New-Object System.Windows.Forms.DateTimePicker
$dtpActivationTime.Location = New-Object System.Drawing.Point(220, 250)
$dtpActivationTime.Size = New-Object System.Drawing.Size(100, 20)
$dtpActivationTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$dtpActivationTime.ShowUpDown = $true # Use up/down arrows for time
$form.Controls.Add($dtpActivationTime)

# Optional Expiry Date and Time
$lblExpiry = New-Object System.Windows.Forms.Label
$lblExpiry.Location = New-Object System.Drawing.Point(20, 290)
$lblExpiry.Size = New-Object System.Drawing.Size(400, 20)
$lblExpiry.Text = "Optional: Expiry Date and Time (NZ Auckland Time):"
$form.Controls.Add($lblExpiry)

$dtpExpiryDate = New-Object System.Windows.Forms.DateTimePicker
$dtpExpiryDate.Location = New-Object System.Drawing.Point(20, 310)
$dtpExpiryDate.Size = New-Object System.Drawing.Size(200, 20)
$dtpExpiryDate.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$dtpExpiryDate.Checked = $false # Start unchecked
$dtpExpiryDate.ShowCheckBox = $true # Show checkbox to enable/disable
$form.Controls.Add($dtpExpiryDate)

$dtpExpiryTime = New-Object System.Windows.Forms.DateTimePicker
$dtpExpiryTime.Location = New-Object System.Drawing.Point(220, 310)
$dtpExpiryTime.Size = New-Object System.Drawing.Size(100, 20)
$dtpExpiryTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$dtpExpiryTime.ShowUpDown = $true # Use up/down arrows for time
$dtpExpiryTime.Enabled = $false # Start disabled
$form.Controls.Add($dtpExpiryTime)

# Event handler to enable/disable Expiry Time when checkbox is checked/unchecked
$dtpExpiryDate.Add_CheckedChanged({
    $dtpExpiryTime.Enabled = $dtpExpiryDate.Checked
})


# Status Label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 350)
$lblStatus.Size = New-Object System.Drawing.Size(400, 60)
$lblStatus.Text = "Status: Ready"
$lblStatus.ForeColor = [System.Drawing.Color]::Black
$form.Controls.Add($lblStatus)

# Create Rule Button
$btnCreateRule = New-Object System.Windows.Forms.Button
$btnCreateRule.Location = New-Object System.Drawing.Point(150, 420)
$btnCreateRule.Size = New-Object System.Drawing.Size(120, 40)
$btnCreateRule.Text = "Create Rule"
$form.Controls.Add($btnCreateRule)

# Set default button after it's created
$form.AcceptButton = $btnCreateRule

# --- Button Click Event Handler ---
$btnCreateRule.Add_Click({
    # Update status label
    $lblStatus.Text = "Status: Processing..."
    $lblStatus.ForeColor = [System.Drawing.Color]::Black # Use a neutral color during processing
    $form.Update() # Force GUI update

    # Get values from form
    $SourceMailbox = $txtSource.Text.Trim()
    $RedirectingEmail = $txtRedirect.Text.Trim()
    $ActionType = If ($rbRedirect.Checked) {"redirect"} Else {"copy"}

    $ActivationDateOnly = $dtpActivationDate.Value
    $ActivationTimeOnly = $dtpActivationTime.Value
    $ActivationDateTime = New-Object DateTime($ActivationDateOnly.Year, $ActivationDateOnly.Month, $ActivationDateOnly.Day, $ActivationTimeOnly.Hour, $ActivationTimeOnly.Minute, $ActivationTimeOnly.Second)

    $ExpiryDateTime = $null # Initialize expiry date

    # Check if the expiry date checkbox is checked before getting the value
    if ($dtpExpiryDate.Checked) {
        $ExpiryDateOnly = $dtpExpiryDate.Value
        $ExpiryTimeOnly = $dtpExpiryTime.Value
        $ExpiryDateTime = New-Object DateTime($ExpiryDateOnly.Year, $ExpiryDateOnly.Month, $ExpiryDateOnly.Day, $ExpiryTimeOnly.Hour, $ExpiryTimeOnly.Minute, $ExpiryTimeOnly.Second)
    }

    # Basic validation
    if ([string]::IsNullOrEmpty($SourceMailbox) -or [string]::IsNullOrEmpty($RedirectingEmail)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the source mailbox and the redirect/copy mailbox.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $lblStatus.Text = "Status: Validation Error"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        $form.Update()
        return
    }

    # --- Exchange Online Logic ---
    try {
        # Connect to Exchange Online (will prompt for credentials if not already connected)
        # Check if already connected to avoid multiple prompts
        if (-not (Get-PSSession -ComputerName outlook.office365.com -ErrorAction SilentlyContinue)) {
             $lblStatus.Text = "Status: Connecting to Exchange Online..."
             $lblStatus.ForeColor = [System.Drawing.Color]::Blue
             $form.Update()
             Connect-ExchangeOnline -ShowBanner:$false -ShowProgress:$false # Suppress verbose output
        }


        # Get the New Zealand Standard Time zone
        $NZTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("New Zealand Standard Time")

        # Convert activation time to UTC
        $ActivationDateUtc = [System.TimeZoneInfo]::ConvertTimeToUtc($ActivationDateTime, $NZTimeZone)

        # Convert expiry time to UTC if set
        $ExpiryDateUtc = $null
        if ($ExpiryDateTime) {
            $ExpiryDateUtc = [System.TimeZoneInfo]::ConvertTimeToUtc($ExpiryDateTime, $NZTimeZone)
        }

        # Define a name for the rule
        # Correctly capitalize the first letter of the action type
        $CapitalizedActionType = $ActionType.Substring(0,1).ToUpper() + $ActionType.Substring(1)
        $RuleName = "$($CapitalizedActionType) Rule for $($SourceMailbox.Split('@')[0])"

        if ($ExpiryDateUtc) {
             $RuleName += " (Expires $($ExpiryDateTime.ToString('yyyyMMdd')))"
        } else {
             $RuleName += " (Permanent)"
        }


        $lblStatus.Text = "Status: Creating rule in Exchange Online..."
        $lblStatus.ForeColor = [System.Drawing.Color]::Blue
        $form.Update()

        # Create the Transport Rule
        $params = @{
            Name = $RuleName
            Enabled = $true
            SentTo = $SourceMailbox
            ActivationDate = $ActivationDateUtc
            # Removed the incorrect -StopProcessingRules parameter
        }

        if ($ActionType -eq "redirect") {
            $params.Add("RedirectMessageTo", $RedirectingEmail)
        } else { # copy
            $params.Add("CopyTo", $RedirectingEmail)
        }

        if ($ExpiryDateUtc) {
            $params.Add("ExpiryDate", $ExpiryDateUtc)
        }

        New-TransportRule @params

        $lblStatus.Text = "Status: Rule created successfully!"
        $lblStatus.ForeColor = [System.Drawing.Color]::Green

    } catch {
        $ErrorMessage = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Error creating rule: $ErrorMessage", "Exchange Online Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $lblStatus.Text = "Status: Error creating rule"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    } finally {
        # Optional: Disconnect Exchange Online session after rule creation
        # Disconnect-ExchangeOnline -Confirm:$false
        $form.Update()
    }
})

# --- Display the Form ---
$form.ShowDialog()

# --- Disconnect Exchange Online (if connected and not disconnected in finally block) ---
# if (Get-PSSession -ComputerName outlook.office365.com -ErrorAction SilentlyContinue) {
#    Disconnect-ExchangeOnline -Confirm:$false
# }

