#region Function Declarations
Function Send-WarningEmail {
    Param(
        [Int]$Days,
        [String]$ExpiryLine,
        [String]$Date,
        [String]$Email
    )

    $body = @"
    <html
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word"
    xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
    xmlns="http://www.w3.org/TR/REC-html40"
    >
    <head>
      <meta http-equiv="Content-Type" content="text/html; charset=us-ascii" />
      <meta name="Generator" content="Microsoft Word 15 (filtered medium)" />
      <style>
        <!--
            /* Font Definitions */
            @font-face
                {font-family:"Cambria Math";
                panose-1:2 4 5 3 5 4 6 3 2 4;}
            @font-face
                {font-family:Calibri;
                panose-1:2 15 5 2 2 2 4 3 2 4;}
            /* Style Definitions */
            p.MsoNormal, li.MsoNormal, div.MsoNormal
                {margin:0in;
                font-size:11.0pt;
                font-family:"Calibri",sans-serif;}
            a:link, span.MsoHyperlink
                {mso-style-priority:99;
                color:#0563C1;
                text-decoration:underline;}
            p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
                {mso-style-priority:34;
                margin-top:0in;
                margin-right:0in;
                margin-bottom:0in;
                margin-left:.5in;
                font-size:11.0pt;
                font-family:"Calibri",sans-serif;}
            span.EmailStyle20
                {mso-style-type:personal-reply;
                font-family:"Arial",sans-serif;
                color:windowtext;}
            .MsoChpDefault
                {mso-style-type:export-only;
                font-size:10.0pt;}
            @page WordSection1
                {size:8.5in 11.0in;
                margin:1.0in 1.0in 1.0in 1.0in;}
            div.WordSection1
                {page:WordSection1;}
            /* List Definitions */
            @list l0
                {mso-list-id:155802099;
                mso-list-template-ids:-12136722;}
            @list l0:level1
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l0:level2
                {mso-level-number-format:bullet;
                mso-level-text:o;
                mso-level-tab-stop:1.0in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:"Courier New";
                mso-bidi-font-family:"Times New Roman";}
            @list l0:level3
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:1.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l0:level4
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:2.0in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l0:level5
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:2.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l0:level6
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:3.0in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l0:level7
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:3.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l0:level8
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:4.0in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l0:level9
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:4.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l1
                {mso-list-id:406734789;
                mso-list-template-ids:951063048;}
            @list l1:level1
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l1:level2
                {mso-level-number-format:bullet;
                mso-level-text:o;
                mso-level-tab-stop:1.0in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:"Courier New";
                mso-bidi-font-family:"Times New Roman";}
            @list l1:level3
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:1.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l1:level4
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:2.0in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l1:level5
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:2.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l1:level6
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:3.0in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l1:level7
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:3.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l1:level8
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:4.0in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            @list l1:level9
                {mso-level-number-format:bullet;
                mso-level-text:\F0B7;
                mso-level-tab-stop:4.5in;
                mso-level-number-position:left;
                text-indent:-.25in;
                mso-ansi-font-size:10.0pt;
                font-family:Symbol;}
            ol
                {margin-bottom:0in;}
            ul
                {margin-bottom:0in;}
            -->
      </style>
      <!--[if gte mso 9
        ]><xml>
          <o:shapedefaults v:ext="edit" spidmax="1026" /> </xml
      ><![endif]-->
      <!--[if gte mso 9
        ]><xml>
          <o:shapelayout v:ext="edit">
            <o:idmap v:ext="edit" data="1" /> </o:shapelayout></xml
      ><![endif]-->
    </head>
    <body
      lang="EN-US"
      link="#0563C1"
      vlink="#954F72"
      style="word-wrap: break-word"
    >
      <div class="WordSection1">
        <table
          class="MsoNormalTable"
          border="0"
          cellspacing="0"
          cellpadding="0"
          style="border-collapse: collapse"
        >
          <tr>
            <td
              width="600"
              valign="top"
              style="width: 6.25in; padding: 0in 5.4pt 0in 5.4pt"
            >
              <p class="MsoNormal" style="margin-bottom: 6pt">
                <b
                  ><span
                    style="
                      font-size: 14pt;
                      font-family: 'Arial', sans-serif;
                      color: #3c9ed1;
                    "
                    >Technology</span
                  ></b
                ><o:p></o:p>
              </p>
              <p class="MsoNormal" style="margin-bottom: 6pt">
                <b
                  ><span style="font-family: 'Arial', sans-serif"
                    >$ExpiryLine<o:p></o:p></span
                ></b>
              </p>
              <p class="MsoNormal" style="margin-bottom: 6pt">
                <span style="font-size: 10pt; font-family: 'Arial', sans-serif"
                  >To better protect Firm and client information, we updated our
                  password guidelines as detailed below. Your password must now
                  contain a minimum of <b>16 characters</b>.<b> </b>You will be
                  required to change it every <b>12 months</b>. &nbsp;<o:p></o:p
                ></span>
              </p>
              <p class="MsoNormal" style="margin-bottom: 6pt">
                <span style="font-size: 10pt; font-family: 'Arial', sans-serif"
                  >To update your password:</span
                ><o:p></o:p>
              </p>
              <p class="MsoNormal" style="margin-bottom: 6pt">
                <b
                  ><span style="font-size: 10pt; font-family: 'Arial', sans-serif"
                    >Step 1:</span
                  ></b
                ><span style="font-size: 10pt; font-family: 'Arial', sans-serif">
                  Log in to a Firm computer and connect to the network </span
                ><o:p></o:p>
              </p>
              <p class="MsoNormal" style="margin-bottom: 6pt">
                <b
                  ><span style="font-size: 10pt; font-family: 'Arial', sans-serif"
                    >Step 2:</span
                  ></b
                ><span style="font-size: 10pt; font-family: 'Arial', sans-serif">
                  Press <b>Ctrl + Alt + Delete</b></span
                ><o:p></o:p>
              </p>
              <p class="MsoNormal" style="margin-bottom: 6pt">
                <b
                  ><span style="font-size: 10pt; font-family: 'Arial', sans-serif"
                    >Step 3:</span
                  ></b
                ><span style="font-size: 10pt; font-family: 'Arial', sans-serif">
                  Select <b>Change a password</b></span
                ><o:p></o:p>
              </p>
              <p class="MsoNormal" style="margin-bottom: 6pt">
                <b
                  ><span style="font-size: 10pt; font-family: 'Arial', sans-serif"
                    >Step 4:</span
                  ></b
                ><span style="font-size: 10pt; font-family: 'Arial', sans-serif">
                  Follow the prompts on the screen. Your new password must meet
                  the updated password guidelines below.</span
                ><o:p></o:p>
              </p>
              <table
                class="MsoNormalTable"
                border="0"
                cellspacing="0"
                cellpadding="0"
                style="border-collapse: collapse"
              >
                <tr>
                  <td
                    width="608"
                    valign="top"
                    style="
                      width: 456.2pt;
                      border: solid windowtext 1pt;
                      background: #fff2cc;
                      padding: 0in 5.4pt 0in 5.4pt;
                    "
                  >
                    <p class="MsoNormal" style="margin-bottom: 6pt">
                      <span style="color: black"
                        ><a
                          href="<Password guidlines Link for your company>"
                          ><b
                            ><span
                              style="
                                font-family: 'Arial', sans-serif;
                                color: #3c9ed1;
                              "
                              >Password Guidelines</span
                            ></b
                          ></a
                        ></span
                      ><o:p></o:p>
                    </p>
                    <p class="MsoNormal" style="margin-bottom: 6pt">
                      <span
                        style="
                          font-size: 10pt;
                          font-family: 'Arial', sans-serif;
                          color: black;
                        "
                        >The Firm requires the use of strong passwords to access
                        all Firm Technology. Password changes are required every
                        12 months.</span
                      ><o:p></o:p>
                    </p>
                    <p class="MsoNormal" style="margin-bottom: 6pt">
                      <span
                        style="
                          font-size: 10pt;
                          font-family: 'Arial', sans-serif;
                          color: black;
                        "
                        >A strong password:</span
                      ><o:p></o:p>
                    </p>
                    <ul style="margin-top: 0in" type="disc">
                      <li
                        class="MsoListParagraph"
                        style="
                          margin-bottom: 6pt;
                          margin-left: 0in;
                          mso-list: l0 level1 lfo3;
                        "
                      >
                        <span
                          style="
                            font-size: 10pt;
                            font-family: 'Arial', sans-serif;
                            color: black;
                          "
                          >Contains a minimum of <b>16 characters</b></span
                        ><o:p></o:p>
                      </li>
                      <li
                        class="MsoListParagraph"
                        style="
                          margin-bottom: 6pt;
                          margin-left: 0in;
                          mso-list: l0 level1 lfo3;
                        "
                      >
                        <span
                          style="
                            font-size: 10pt;
                            font-family: 'Arial', sans-serif;
                            color: black;
                          "
                          >Contains a combination of the following categories of
                          characters:</span
                        ><span style="color: black"> </span><o:p></o:p>
                      </li>
                      <ul style="margin-top: 0in" type="circle">
                        <li
                          class="MsoListParagraph"
                          style="
                            margin-bottom: 6pt;
                            margin-left: 0in;
                            mso-list: l0 level2 lfo3;
                          "
                        >
                          <span
                            style="
                              font-size: 10pt;
                              font-family: 'Arial', sans-serif;
                              color: black;
                            "
                            >Uppercase letters</span
                          ><o:p></o:p>
                        </li>
                        <li
                          class="MsoListParagraph"
                          style="
                            margin-bottom: 6pt;
                            margin-left: 0in;
                            mso-list: l0 level2 lfo3;
                          "
                        >
                          <span
                            style="
                              font-size: 10pt;
                              font-family: 'Arial', sans-serif;
                              color: black;
                            "
                            >Lowercase letters</span
                          ><o:p></o:p>
                        </li>
                        <li
                          class="MsoListParagraph"
                          style="
                            margin-bottom: 6pt;
                            margin-left: 0in;
                            mso-list: l0 level2 lfo3;
                          "
                        >
                          <span
                            style="
                              font-size: 10pt;
                              font-family: 'Arial', sans-serif;
                              color: black;
                            "
                            >Numbers</span
                          ><o:p></o:p>
                        </li>
                        <li
                          class="MsoListParagraph"
                          style="
                            margin-bottom: 6pt;
                            margin-left: 0in;
                            mso-list: l0 level2 lfo3;
                          "
                        >
                          <span
                            style="
                              font-size: 10pt;
                              font-family: 'Arial', sans-serif;
                              color: black;
                            "
                            >Special symbols (! @ # $)</span
                          ><o:p></o:p>
                        </li>
                      </ul>
                      <li
                        class="MsoListParagraph"
                        style="
                          margin-bottom: 6pt;
                          margin-left: 0in;
                          mso-list: l0 level1 lfo3;
                        "
                      >
                        <span
                          style="
                            font-size: 10pt;
                            font-family: 'Arial', sans-serif;
                            color: black;
                          "
                          >Is different from your previous three passwords</span
                        ><o:p></o:p>
                      </li>
                      <li
                        class="MsoListParagraph"
                        style="
                          margin-bottom: 6pt;
                          margin-left: 0in;
                          mso-list: l0 level1 lfo3;
                        "
                      >
                        <span
                          style="
                            font-size: 10pt;
                            font-family: 'Arial', sans-serif;
                            color: black;
                          "
                          >Not used for personal accounts</span
                        ><o:p></o:p>
                      </li>
                    </ul>
                    <p class="MsoNormal" style="margin-bottom: 6pt">
                      <b
                        ><span
                          style="
                            font-size: 10pt;
                            font-family: 'Arial', sans-serif;
                            color: black;
                          "
                          >Recommendation: Create a passphrase (e.g. Volleyball is
                          9 to 5!). The use of a space is allowed, but does not
                          count toward length or special symbol
                          requirements.</span
                        ></b
                      ><o:p></o:p>
                    </p>
                  </td>
                </tr>
              </table>
              <p class="MsoNormal" style="margin-bottom: 12pt">
                <b><br /></b
                ><strong
                  ><span style="font-size: 10pt; font-family: 'Arial', sans-serif"
                    >Service Desk</span
                  ></strong
                ><span style="font-size: 10pt; font-family: 'Arial', sans-serif"
                  ><br /></span
                ><strong
                  ><span style="font-size: 9pt; font-family: 'Arial', sans-serif"
                    >T</span
                  ></strong
                ><span style="font-size: 9pt; font-family: 'Arial', sans-serif">
                  +<Service Desk Phone Number>&nbsp;&nbsp;&nbsp;&nbsp;<strong
                    ><span style="font-family: 'Arial', sans-serif"
                      >E</span
                    ></strong
                  > </span
                ><a href="mailto:servicedesk@company.com"
                  ><span style="font-size: 9pt; font-family: 'Arial', sans-serif"
                    >servicedesk@company.com</span
                  ></a
                ><span style="font-size: 9pt; font-family: 'Arial', sans-serif"
                  ><br /><strong
                    ><span style="font-family: 'Arial', sans-serif"
                      >Short Dial</span
                    ></strong
                  >
                  <short dial code></span
                ><o:p></o:p>
              </p>
            </td>
          </tr>
        </table>
       </div>
        <p class="MsoNormal"><o:p>&nbsp;</o:p></p>
      </div>
    </body>
    </html>
"@

    $Splat = @{
        To         = $Email
        From       = 'servicedesk@mail.com.com'
        Body       = $Body
        Subject    =  "Your Password will expire in $Days days on $(Get-Date($Date) -f 'dddd, d MMMM yyyy')"
        SmtpServer = '<smtp_server>'
        BodyAsHtml = $True
    }

    Send-MailMessage @Splat
}
#endregion

#region Get and Trim the list of users.
Try {
    $Users = Get-ADUser -Filter 'employeeid -like "*"' -Properties samaccountname, emailaddress, PassWordLastSet, Enabled, employeeid, created -ErrorAction Stop | Where-Object { $_.enabled -eq $True -and $_.emailaddress -ne $null }
}
Catch [System.Exception] {
    Write-Error "Unable to find any Users. This script has failed."
    $LastExitCode = 1
    exit;
}
#endregion

#region Loop Through Users
$EmailedUsers = New-Object -TypeName System.Collections.ArrayList

Foreach ($User in $Users) {
    #Get Today
    $Today = Get-Date
    #Get Password Last Set
    If ($Null -eq $User.PasswordLastSet) {
        $End = $(Get-Date($User.Created)).AddDays(365)
        [INT]$Timespan = (New-TimeSpan -Start $Today -End $End).days
    }
    Else {
        $End = $(Get-Date($User.PasswordLastSet)).AddDays(365)
        [INT]$Timespan = (New-TimeSpan -Start $Today -End $End).days
    }

    switch ($TimeSpan) {
        14 {
            Write-Host "Sending email to $($User.emailaddress). Password expires on $End"
            $ExpiryLine = "Your Firm password will expire in 14 days."
            #Send-WarningEmail -Days 14 -ExpiryLine $ExpiryLine -Email $User.EmailAddress -Date $End
            $EmailedUsers.add($User) | Out-Null
        }

        7 {
            Write-Host "Sending email to $($User.emailaddress). Password expires on $End"
            $ExpiryLine = "Your Firm password will expire in 7 days."
            #Send-WarningEmail -Days 7 -ExpiryLine $ExpiryLine -Email $User.EmailAddress -Date $End
            $EmailedUsers.add($User) | Out-Null
        }

        2 {
            Write-Host "Sending email to $($User.emailaddress). Password expires on $End"
            $ExpiryLine = "Your Firm password will expire in 2 days."
            #Send-WarningEmail -Days 2 -ExpiryLine $ExpiryLine -Email $User.EmailAddress -Date $End
            $EmailedUsers.add($User) | Out-Null
        }

        1 {
            Write-Host "Sending email to $($User.emailaddress). Password expires on $End"
            $ExpiryLine = "Your Firm password will expire in 1 day."
            #Send-WarningEmail -Days 1 -ExpiryLine $ExpiryLine -Email $User.EmailAddress -Date $End
            $EmailedUsers.add($User) | Out-Null
        }

        0 {
            Write-Host "Sending email to $($User.emailaddress). Password expires on $End"
            $ExpiryLine = "Your Firm password will expire today."
            #Send-WarningEmail -Days 0 -ExpiryLine $ExpiryLine -Email $User.EmailAddress -Date $End
            $EmailedUsers.add($User) | Out-Null
        }
    }
}
#endregion

$Today = Get-Date -f 'd-MMMM-yyyy'
$Path = "$Env:workspace\$($Today).csv"
$EmailedUsers | Select-Object samaccountname, emailaddress, PassWordLastSet | Export-Csv -Path $Path

$Splat = @{
    To          = 'you@company.com'
    From        = 'Jenkins@company.com'
    Body        = "Attached is the list of users who were emailed today."
    Subject     =  "Expiring Accounts Results"
    SmtpServer  = '<smtpserver>'
    BodyAsHtml  = $True
    Attachments = $Path
}

#Send-MailMessage @Splat
