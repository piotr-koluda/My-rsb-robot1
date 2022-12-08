*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the ordered robot.
...                 Embeds the screenshot of the robot to the PDF receipt.
...                 Creates ZIP archive of the receipts and the images.

Library             RPA.Browser.Selenium    auto_close=${False}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Tables
Library             RPA.PDF
Library             RPA.Desktop
Library             OperatingSystem
Library             RPA.FileSystem
Library             RPA.Email.Exchange
Library             RPA.Archive
Library             RPA.Dialogs
Library             RPA.Robocorp.Vault


*** Variables ***
${GLOBAL_RETRY_AMOUNT}          3x
${GLOBAL_RETRY_INTERVAL}        0.5s
${GLOBAL_DESTINATION_PATH}      ${OUTPUT_DIR}${/}Temp_folder


*** Tasks ***
Minimal task
    ${URL_to_Page} =    Get Secret    Order URL

    # Prepare Dialog window to get user data
    Add text input    name=CSV    label=Location CSV file
    ${url_to_check} =    Run dialog    title=Input from
    # Create temp folder to get data
    Create_Folder_Temp

    Launch_Browser    ${URL_to_Page}[address]
    Navigate_To_Form
    Get_CSV_File    ${url_to_check.CSV}
    ${cases} =
    ...    Read table from CSV
    ...    orders.csv
    ...    header= True

    FOR    ${cases}    IN    @{cases}
        Log    ${cases}[Order number]
        Fill_Form    ${cases}
        Remove_Files    ${cases}[Order number]
    END
    ${zip_folder_name} =    Set Variable    ${OUTPUT_DIR}${/}PDFs_Zip.zip

    Archive Folder With Zip
    ...    ${GLOBAL_DESTINATION_PATH}
    ...    ${zip_folder_name}
    Log    Done.
    [Teardown]    Close Browser


*** Keywords ***
Create_Folder_Temp
    RPA.FileSystem.Create Directory    ${GLOBAL_DESTINATION_PATH}

Get_CSV_File
    [Arguments]    ${URL_to_download}
    Download    ${URL_to_download}    overwrite=True

Navigate_To_Form
    Click Link    alias:Link_OrderRobot
    Wait Until Element is Visible    alias:CSS_ModalHeader
    Click Button When Visible    xpath://button[contains(.,'OK')]

Launch_Browser
    [Arguments]    ${URL}
    Open Available Browser    ${URL}    browser_selection=chrome
    Maximize Browser Window

Check_Error
    ${rep} =    Is Element Visible    xpath://div[@id='receipt']

    WHILE    ${rep} == False
        Wait Until Keyword Succeeds
        ...    ${GLOBAL_RETRY_AMOUNT}
        ...    ${GLOBAL_RETRY_INTERVAL}
        ...    Wait And Click Button    //button[@id="order"]
        ${rep} =    Is Element Visible    xpath://div[@id='receipt']
    END

Save_PDF_1
    [Arguments]    ${Order number}
    Wait Until Element Is Visible    xpath://div[@id='receipt']
    ${sales_results_html} =    Get Element Attribute    xpath://div[@id='receipt']    outerHTML
    Html To Pdf    ${sales_results_html}    ${GLOBAL_DESTINATION_PATH}${/}${Order number}.pdf

Close_Popup
    Click Button When Visible    id:order-another
    Wait Until Element is Visible    alias:CSS_ModalHeader
    Click Button When Visible    xpath://button[contains(.,'OK')]

Fill_Form
    [Arguments]    ${cases}

    Select From List By Index    alias:CSS_head    ${cases}[Head]
    Select Radio Button    body    ${cases}[Body]
    Input Text    alias:xpath_legs    ${cases}[Legs]
    Input Text    alias:xpath_address    ${cases}[Address]
    Click Button    id:preview
    Click Button    id:order

    TRY
        Save_PDF_1    ${cases}[Order number]
    EXCEPT
        Check_Error
        Save_PDF_1    ${cases}[Order number]
    FINALLY
        ${screenshot} =    Take_A_Screenshot    ${cases}[Order number]
        Combine_pdf
        ...    ${GLOBAL_DESTINATION_PATH}${/}${cases}[Order number].pdf
        ...    ${screenshot}
        ...    ${cases}[Order number]
        Close_Popup
    END

Take_A_Screenshot
    [Arguments]    ${order}
    Screenshot    xpath://div[@id='robot-preview-image']    ${GLOBAL_DESTINATION_PATH}${/}${order}.png
    RETURN    ${GLOBAL_DESTINATION_PATH}${/}${order}.png

Combine_pdf
    [Arguments]    ${pdf_Path}    ${jpg_path}    ${order_no}
    Open Pdf    ${pdf_Path}
    @{file_list} =    Create List
    ...    ${jpg_path}
    ...    ${pdf_Path}
    Add Files To Pdf    ${file_list}    ${GLOBAL_DESTINATION_PATH}${/}Order_No_${order_no}.pdf
    Close All Pdfs

Remove_Files
    [Arguments]    ${file_number}

    ${isPdfFile} =    Does File Exist    ${GLOBAL_DESTINATION_PATH}${/}${file_number}.pdf
    IF    ${isPdfFile} == True
        OperatingSystem.Remove File    ${GLOBAL_DESTINATION_PATH}${/}${file_number}.pdf
    END

    ${isPNGfile} =    Does File Exist    ${GLOBAL_DESTINATION_PATH}${/}${file_number}.png
    IF    ${isPNGfile} == True
        OperatingSystem.Remove File    ${GLOBAL_DESTINATION_PATH}${/}${file_number}.png
    END
