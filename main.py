import pywifi
import time
import csv
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Create a WiFi interface scanner
wifi = pywifi.PyWiFi()

# Create an interface object
iface = wifi.interfaces()[0]  # You may need to change the index if you have multiple interfaces

# Scan for available Wi-Fi networks
iface.scan()
time.sleep(2)  # Wait for a moment to ensure the scan is complete

# Get the scan results
scan_results = iface.scan_results()

# Define the name of the output text file and CSV file
output_txt_file = "wifi_networks.txt"
output_csv_file = "wifi_networks.csv"

# Create a set to keep track of unique SSIDs
unique_ssids = set()

# Open the text file for writing with UTF-8 encoding
with open(output_txt_file, "w", encoding="utf-8") as txt_file:
    txt_file.write("List of available Wi-Fi networks:\n")

    # Open the CSV file for writing
    with open(output_csv_file, "w", newline="", encoding="utf-8") as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(["SSID", "Signal Strength (dBm)", "BSSID (MAC Address)", "Security"])

        # Write each network to the text file and CSV file, removing duplicates
        for wifi_network in scan_results:
            ssid = wifi_network.ssid
            signal_strength = wifi_network.signal
            bssid = wifi_network.bssid
            security = wifi_network.akm[0]

            # Check if this SSID is unique
            if ssid not in unique_ssids:
                txt_file.write(
                    f"SSID: {ssid}, Signal Strength: {signal_strength} dBm, BSSID: {bssid}, Security: {security} \n")
                csv_writer.writerow([ssid, signal_strength, bssid, security])
                unique_ssids.add(ssid)

print(f"Available Wi-Fi networks saved to {output_txt_file} and {output_csv_file}")

# Email configuration
smtp_server = '172.24.46.52'
smtp_port = 25
sender_email = 'it.support.dnph.a1b@ap.denso.com'
cc_emails = ['philip.lominoque.a5d@ap.denso.com']  # Add your CC recipients

current_timestamp = datetime.now()
cur_month = current_timestamp.strftime("%B")
cur_date = current_timestamp.strftime("%d")
cur_year = current_timestamp.strftime("%Y")
# Create the email message
subject = "[Audit Purpose] Open WiFi Reminder! as of " + str(cur_month) + " " + str(cur_date) + " " + str(cur_year)
text = "Kindly see the attachment regarding to the open network. Thank you!"
message = """\
<!DOCTYPE html>
<html
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  lang="en"
>
  <head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="viewport" content="width=device-width,initial-scale=1" />
    <!--[if mso
      ]><xml
        ><o:OfficeDocumentSettings
          ><o:PixelsPerInch>96</o:PixelsPerInch
          ><o:AllowPNG /></o:OfficeDocumentSettings></xml
    ><![endif]-->
    <style>
      * {
        box-sizing: border-box;
      }
      body {
        margin: 0;
        padding: 0;
      }
      a[x-apple-data-detectors] {
        color: inherit !important;
        text-decoration: inherit !important;
      }
      #MessageViewBody a {
        color: inherit;
        text-decoration: none;
      }
      p {
        line-height: inherit;
      }
      .desktop_hide,
      .desktop_hide table {
        mso-hide: all;
        display: none;
        max-height: 0;
        overflow: hidden;
      }
      .image_block img + div {
        display: none;
      }
      @media (max-width: 690px) {
        .row-content {
          width: 100% !important;
        }
        .mobile_hide {
          display: none;
        }
        .stack .column {
          width: 100%;
          display: block;
        }
        .mobile_hide {
          min-height: 0;
          max-height: 0;
          max-width: 0;
          overflow: hidden;
          font-size: 0;
        }
        .desktop_hide,
        .desktop_hide table {
          display: table !important;
          max-height: none !important;
        }
      }
    </style>
  </head>
  <body
    style="
      background-color: #d1e3f0;
      margin: 0;
      padding: 0;
      -webkit-text-size-adjust: none;
      text-size-adjust: none;
    "
  >
    <table
      class="nl-container"
      width="100%"
      border="0"
      cellpadding="0"
      cellspacing="0"
      role="presentation"
      style="
        mso-table-lspace: 0;
        mso-table-rspace: 0;
        background-color: #d1e3f0;
      "
    >
      <tbody>
        <tr>
          <td>
            <table
              class="row row-1"
              align="center"
              width="100%"
              border="0"
              cellpadding="0"
              cellspacing="0"
              role="presentation"
              style="mso-table-lspace: 0; mso-table-rspace: 0"
            >
              <tbody>
                <tr>
                  <td>
                    <table
                      class="row-content stack"
                      align="center"
                      border="0"
                      cellpadding="0"
                      cellspacing="0"
                      role="presentation"
                      style="
                        mso-table-lspace: 0;
                        mso-table-rspace: 0;
                        color: #000;
                        width: 670px;
                      "
                      width="670"
                    >
                      <tbody>
                        <tr>
                          <td
                            class="column column-1"
                            width="100%"
                            style="
                              mso-table-lspace: 0;
                              mso-table-rspace: 0;
                              font-weight: 400;
                              text-align: left;
                              vertical-align: top;
                              border-top: 0;
                              border-right: 0;
                              border-bottom: 0;
                              border-left: 0;
                            "
                          >
                            <table
                              class="empty_block block-1"
                              width="100%"
                              border="0"
                              cellpadding="0"
                              cellspacing="0"
                              role="presentation"
                              style="mso-table-lspace: 0; mso-table-rspace: 0"
                            >
                              <tr>
                                <td class="pad"><div></div></td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
            <table
              class="row row-2"
              align="center"
              width="100%"
              border="0"
              cellpadding="0"
              cellspacing="0"
              role="presentation"
              style="mso-table-lspace: 0; mso-table-rspace: 0"
            >
              <tbody>
                <tr>
                  <td>
                    <table
                      class="row-content stack"
                      align="center"
                      border="0"
                      cellpadding="0"
                      cellspacing="0"
                      role="presentation"
                      style="
                        mso-table-lspace: 0;
                        mso-table-rspace: 0;
                        background-color: #fff;
                        color: #000;
                        width: 670px;
                      "
                      width="670"
                    >
                      <tbody>
                        <tr>
                          <td
                            class="column column-1"
                            width="100%"
                            style="
                              mso-table-lspace: 0;
                              mso-table-rspace: 0;
                              font-weight: 400;
                              text-align: left;
                              border-top: 4px solid silver;
                              padding-bottom: 25px;
                              padding-left: 15px;
                              padding-right: 15px;
                              padding-top: 25px;
                              vertical-align: top;
                              border-right: 0;
                              border-bottom: 0;
                              border-left: 0;
                            "
                          >
                            <table
                              class="text_block block-1"
                              width="100%"
                              border="0"
                              cellpadding="10"
                              cellspacing="0"
                              role="presentation"
                              style="
                                mso-table-lspace: 0;
                                mso-table-rspace: 0;
                                word-break: break-word;
                              "
                            >
                              <tr>
                                <td class="pad">
                                  <div style="font-family: sans-serif">
                                    <div
                                      class
                                      style="
                                        font-size: 12px;
                                        font-family: Verdana, Geneva, sans-serif;
                                        mso-line-height-alt: 18px;
                                        color: #171717;
                                        line-height: 1.5;
                                      "
                                    >
                                      <p
                                        style="
                                          margin: 0;
                                          font-size: 12px;
                                          text-align: center;
                                          mso-line-height-alt: 36px;
                                        "
                                      >
                                        <span style="font-size: 24px"
                                          ><strong
                                            >Kindly see the attachment regarding to the open network. Thank you and have a nice day! </strong
                                          ></span
                                        >
                                      </p>

                                    </div>
                                  </div>
                                </td>
                              </tr>
                            </table>
                            <table
                              class="button_block block-2"
                              width="100%"
                              border="0"
                              cellpadding="10"
                              cellspacing="0"
                              role="presentation"
                              style="mso-table-lspace: 0; mso-table-rspace: 0"
                            >
                              <tr>
                                <td class="pad">

                                </td>
                              </tr>
                            </table>
                            <table
                              class="text_block block-3"
                              width="100%"
                              border="0"
                              cellpadding="10"
                              cellspacing="0"
                              role="presentation"
                              style="
                                mso-table-lspace: 0;
                                mso-table-rspace: 0;
                                word-break: break-word;
                              "
                            >
                              <tr>
                                <td class="pad">
                                  <div style="font-family: sans-serif">

                                  <div
                                      class
                                      style="
                                        font-size: 12px;
                                        font-family: Verdana, Geneva, sans-serif;
                                        mso-line-height-alt: 18px;
                                        color: #595959;
                                        line-height: 1.5;
                                      "
                                    >
                                      <p
                                        style="
                                          margin: 0;
                                          text-align: center;
                                          mso-line-height-alt: 27px;">
                                      </p>
                                    </div>

                                    <div
                                      class
                                      style="
                                        font-size: 12px;
                                        font-family: Verdana, Geneva, sans-serif;
                                        mso-line-height-alt: 18px;
                                        color: #595959;
                                        line-height: 1.5;
                                      "
                                    >
                                      <p
                                        style="
                                          margin: 0;
                                          text-align: center;
                                          mso-line-height-alt: 27px;">
                                      </p>
                                    </div>

                                    <div
                                      class
                                      style="
                                        font-size: 12px;
                                        font-family: Verdana, Geneva, sans-serif;
                                        mso-line-height-alt: 18px;
                                        color: #595959;
                                        line-height: 1.5;
                                      "
                                    >
                                      <p
                                        style="
                                          margin: 0;
                                          text-align: center;
                                          mso-line-height-alt: 27px;">
                                      </p>
                                    </div>

                                    <div
                                      class
                                      style="
                                        font-size: 12px;
                                        font-family: Verdana, Geneva, sans-serif;
                                        mso-line-height-alt: 18px;
                                        color: #595959;
                                        line-height: 1.5;
                                      "
                                    >
                                      <p
                                        style="
                                          margin: 0;
                                          text-align: center;
                                          mso-line-height-alt: 27px;">
                                      </p>
                                    </div>

                                    <div
                                      class
                                      style="
                                        font-size: 12px;
                                        font-family: Verdana, Geneva, sans-serif;
                                        mso-line-height-alt: 18px;
                                        color: #595959;
                                        line-height: 1.5;
                                      "
                                    >
                                      <p
                                        style="
                                          margin: 0;
                                          text-align: center;
                                          mso-line-height-alt: 27px;">
                                      </p>
                                    </div>


                                  </div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
            <table
              class="row row-3"
              align="center"
              width="100%"
              border="0"
              cellpadding="0"
              cellspacing="0"
              role="presentation"
              style="mso-table-lspace: 0; mso-table-rspace: 0"
            >
              <tbody>
                <tr>
                  <td>
                    <table
                      class="row-content stack"
                      align="center"
                      border="0"
                      cellpadding="0"
                      cellspacing="0"
                      role="presentation"
                      style="
                        mso-table-lspace: 0;
                        mso-table-rspace: 0;
                        background-color: #ececec;
                        color: #000;
                        width: 670px;
                      "
                      width="670"
                    >
                      <tbody>
                        <tr>
                          <td
                            class="column column-1"
                            width="100%"
                            style="
                              mso-table-lspace: 0;
                              mso-table-rspace: 0;
                              font-weight: 400;
                              text-align: left;
                              padding-bottom: 5px;
                              padding-top: 5px;
                              vertical-align: top;
                              border-top: 0;
                              border-right: 0;
                              border-bottom: 0;
                              border-left: 0;
                            "
                          >
                            <table
                              class="text_block block-1"
                              width="100%"
                              border="0"
                              cellpadding="10"
                              cellspacing="0"
                              role="presentation"
                              style="
                                mso-table-lspace: 0;
                                mso-table-rspace: 0;
                                word-break: break-word;
                              "
                            >
                              <tr>
                                <td class="pad">
                                  <div style="font-family: sans-serif">
                                    <div
                                      class
                                      style="
                                        font-size: 12px;
                                        font-family: Verdana, Geneva, sans-serif;
                                        mso-line-height-alt: 18px;
                                        color: #595959;
                                        line-height: 1.5;
                                      "
                                    >
                                      <p
                                        style="
                                          margin: 0;
                                          font-size: 14px;
                                          text-align: center;
                                          mso-line-height-alt: 18px;
                                        "
                                      >

                                      </p>
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
            <table
              class="row row-4"
              align="center"
              width="100%"
              border="0"
              cellpadding="0"
              cellspacing="0"
              role="presentation"
              style="mso-table-lspace: 0; mso-table-rspace: 0"
            >
              <tbody>
                <tr>
                  <td>
                    <table
                      class="row-content stack"
                      align="center"
                      border="0"
                      cellpadding="0"
                      cellspacing="0"
                      role="presentation"
                      style="
                        mso-table-lspace: 0;
                        mso-table-rspace: 0;
                        background-color: #ececec;
                        color: #333;
                        width: 670px;
                      "
                      width="670"
                    >
                      <tbody>
                        <tr>
                          <td
                            class="column column-1"
                            width="50%"
                            style="
                              mso-table-lspace: 0;
                              mso-table-rspace: 0;
                              font-weight: 400;
                              text-align: left;
                              padding-left: 10px;
                              padding-right: 10px;
                              padding-top: 10px;
                              vertical-align: top;
                              border-top: 0;
                              border-right: 0;
                              border-bottom: 0;
                              border-left: 0;
                            "
                          >
                            <table
                              class="empty_block block-1"
                              width="100%"
                              border="0"
                              cellpadding="0"
                              cellspacing="0"
                              role="presentation"
                              style="mso-table-lspace: 0; mso-table-rspace: 0"
                            >
                              <tr>
                                <td class="pad"><div></div></td>
                              </tr>
                            </table>
                          </td>
                          <td
                            class="column column-2"
                            width="50%"
                            style="
                              mso-table-lspace: 0;
                              mso-table-rspace: 0;
                              font-weight: 400;
                              text-align: left;
                              padding-bottom: 10px;
                              padding-left: 10px;
                              padding-right: 10px;
                              padding-top: 10px;
                              vertical-align: top;
                              border-top: 0;
                              border-right: 0;
                              border-bottom: 0;
                              border-left: 0;
                            "
                          >
                            <table
                              class="empty_block block-1"
                              width="100%"
                              border="0"
                              cellpadding="0"
                              cellspacing="0"
                              role="presentation"
                              style="mso-table-lspace: 0; mso-table-rspace: 0"
                            >
                              <tr>
                                <td class="pad"><div></div></td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
            <table
              class="row row-5"
              align="center"
              width="100%"
              border="0"
              cellpadding="0"
              cellspacing="0"
              role="presentation"
              style="mso-table-lspace: 0; mso-table-rspace: 0"
            >
              <tbody>
                <tr>
                  <td>
                    <table
                      class="row-content stack"
                      align="center"
                      border="0"
                      cellpadding="0"
                      cellspacing="0"
                      role="presentation"
                      style="
                        mso-table-lspace: 0;
                        mso-table-rspace: 0;
                        color: #000;
                        width: 670px;
                      "
                      width="670"
                    >
                      <tbody>
                        <tr>
                          <td
                            class="column column-1"
                            width="100%"
                            style="
                              mso-table-lspace: 0;
                              mso-table-rspace: 0;
                              font-weight: 400;
                              text-align: left;
                              padding-bottom: 5px;
                              padding-top: 5px;
                              vertical-align: top;
                              border-top: 0;
                              border-right: 0;
                              border-bottom: 0;
                              border-left: 0;
                            "
                          >
                            <table
                              class="empty_block block-1"
                              width="100%"
                              border="0"
                              cellpadding="0"
                              cellspacing="0"
                              role="presentation"
                              style="mso-table-lspace: 0; mso-table-rspace: 0"
                            >
                              <tr>
                                <td class="pad"><div></div></td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
            <table
              class="row row-6"
              align="center"
              width="100%"
              border="0"
              cellpadding="0"
              cellspacing="0"
              role="presentation"
              style="mso-table-lspace: 0; mso-table-rspace: 0"
            >
              <tbody>
                <tr>
                  <td>
                    <table
                      class="row-content stack"
                      align="center"
                      border="0"
                      cellpadding="0"
                      cellspacing="0"
                      role="presentation"
                      style="
                        mso-table-lspace: 0;
                        mso-table-rspace: 0;
                        color: #000;
                        width: 670px;
                      "
                      width="670"
                    >
                      <tbody>
                        <tr>
                          <td
                            class="column column-1"
                            width="100%"
                            style="
                              mso-table-lspace: 0;
                              mso-table-rspace: 0;
                              font-weight: 400;
                              text-align: left;
                              padding-bottom: 30px;
                              padding-top: 30px;
                              vertical-align: top;
                              border-top: 0;
                              border-right: 0;
                              border-bottom: 0;
                              border-left: 0;
                            "
                          >
                            <table
                              class="html_block block-1"
                              width="100%"
                              border="0"
                              cellpadding="0"
                              cellspacing="0"
                              role="presentation"
                              style="mso-table-lspace: 0; mso-table-rspace: 0"
                            >
                              <tr>
                                <td class="pad">
                                  <div
                                    style="
                                      font-family: Verdana, Geneva, sans-serif;
                                      text-align: center;
                                    "
                                    align="center"
                                  ></div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>
    <!-- End -->
    <div style="background-color: transparent">
      <div
        style="
          margin: 0 auto;
          min-width: 320px;
          max-width: 500px;
          overflow-wrap: break-word;
          word-wrap: break-word;
          word-break: break-word;
          background-color: transparent;
        "
        class="block-grid"
      >
        <div
          style="
            border-collapse: collapse;
            display: table;
            width: 100%;
            background-color: transparent;
          "
        >
          <!--[if (mso)|(IE)]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:transparent;" align="center"><table cellpadding="0" cellspacing="0" border="0" style="width: 500px;"><tr class="layout-full-width" style="background-color:transparent;"><![endif]-->
          <!--[if (mso)|(IE)]><td align="center" width="500" style=" width:500px; padding-right: 0px; padding-left: 0px; padding-top:15px; padding-bottom:15px; border-top: 0px solid transparent; border-left: 0px solid transparent; border-bottom: 0px solid transparent; border-right: 0px solid transparent;" valign="top"><![endif]-->
          <div
            class="col num12"
            style="
              min-width: 320px;
              max-width: 500px;
              display: table-cell;
              vertical-align: top;
            "
          >
            <div style="background-color: transparent; width: 100% !important">
              <!--[if (!mso)&(!IE)]><!-->
              <div
                style="
                  border-top: 0px solid transparent;
                  border-left: 0px solid transparent;
                  border-bottom: 0px solid transparent;
                  border-right: 0px solid transparent;
                  padding-top: 15px;
                  padding-bottom: 15px;
                  padding-right: 0px;
                  padding-left: 0px;
                "
              >
                <!--<![endif]-->

                <div
                  align="center"
                  class="img-container center autowidth"
                  style="padding-right: 0px; padding-left: 0px"
                >
                  <!--[if mso]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="padding-right: 0px; padding-left: 0px;" align="center"><![endif]-->


                  <!--[if mso]></td></tr></table><![endif]-->
                </div>

                <!--[if (!mso)&(!IE)]><!-->
              </div>
              <!--<![endif]-->
            </div>
          </div>
          <!--[if (mso)|(IE)]></td></tr></table></td></tr></table><![endif]-->
        </div>
      </div>
    </div>
  </body>
</html>
"""
msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = sender_email
msg['Cc'] = ', '.join(cc_emails)
msg['Subject'] = subject
msg.attach(MIMEText(message, 'html'))
#msg.attach(MIMEText(message, 'plain'))

# Attach the CSV file
with open(output_csv_file, 'rb') as csv_attachment:
    part = MIMEApplication(csv_attachment.read(), Name=output_csv_file)
part['Content-Disposition'] = f'attachment; filename="{output_csv_file}"'
msg.attach(part)

# Connect to the SMTP server and send the email
try:
    server = smtplib.SMTP(smtp_server, smtp_port)
    to_emails = [sender_email] + cc_emails
    server.sendmail(sender_email, to_emails, msg.as_string())
    server.quit()
    print("Email sent successfully!")
except Exception as e:
    print(f"Email sending failed: {str(e)}")