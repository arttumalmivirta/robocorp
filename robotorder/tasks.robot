*** Settings ***
Library     RPA.Browser.Selenium
Library     RPA.HTTP
Library     RPA.Tables
Library     RPA.Robocorp.WorkItems
Library     Collections
Library     RPA.PDF
Library     RPA.Windows
Library     RPA.FileSystem
Library     RPA.Archive


*** Variables ***
${robotordersurl}       https://robotsparebinindustries.com/#/robot-order
${ordersfileurl}        https://robotsparebinindustries.com/orders.csv
${orderscsvfilepath}    C:/Users/ArttuMalmivirta/OneDrive - Digital Workforce/Robocorp/robotorder/orders.csv
${outputfolder}         C:/Users/ArttuMalmivirta/OneDrive - Digital Workforce/Robocorp/robotorder/output/
${robotroot}            C:/Users/ArttuMalmivirta/OneDrive - Digital Workforce/Robocorp/robotorder/


*** Tasks ***
Robot Orders
    Launch Browser    ${robotordersurl}
    Handle All Orders    ${orderscsvfilepath}    orderscsvurl=${ordersfileurl}    pdffolder=${pdffolder}
    Archive Pdfs
    Close Orders Browser


*** Keywords ***
Launch Browser
    [Arguments]    ${url}=url
    Open Available Browser    ${url}    maximized=${TRUE}
    Wait Until Element Is Visible    alias:Modalheader    5
    Click Button    OK
    Wait Until Element Is Not Visible    alias:Modalheader    30

Download Orders File
    [Arguments]    ${ordersurl}=ordersurl    ${targetfilepath}=targetfilepath
    Download    ${ordersurl}    ${targetfilepath}    overwrite=${TRUE}
    ${orders}=    Read table from CSV    ${targetfilepath}
    ${orders}=    Set Variable

Read Model Name
    [Arguments]    ${partnumber}=partnumber
    ${issmodelinfovisible}=    Is Element Visible    id:model-info
    IF    $issmodelinfovisible is $false    Click Button    Show model info
    Wait Until Element Is Visible    id:model-info    10
    ${partnumberint}=    Convert To Integer    ${partnumber}
    ${partnumberint}=    Set Variable    ${${partnumberint}+1}
    ${modelname}=    RPA.Browser.Selenium.Get Table Cell    id:model-info    ${partnumberint}    1
    RETURN    ${modelname}

Close Orders Browser
    Close Browser

Build Robot
    [Arguments]    ${headnumber}=headnumber    ${bodynumber}=bodynumber    ${legsnumber}=legsnumber    ${address}=address
    Wait Until Element Is Visible    id:head    10
    ${isrightsvisible}=    Is Element Visible    alias:Modalheader
    IF    $isrightsvisible    Click Button    OK
    ${modelname}=    Read Model Name    ${headnumber}
    ${headlist}=    Get List Items    id:head    ${FALSE}
    Log List    ${headlist}
    Select From List By Label    id:head    ${modelname} head
    Select Radio Button    body    ${bodynumber}
    Input Text    xpath://html/body/div/div/div[1]/div/div[1]/form/div[3]/input    ${legsnumber}
    Input Text    id:address    ${address}
    Click Button    id:preview
    Wait Until Element Is Visible    id:robot-preview-image    5
    ${counter}=    Set Variable    0
    ${counterint}=    Convert To Integer    ${counter}
    WHILE    $counterint < 5
        ${counterint}=    Set Variable    ${${counterint}+1}
        Click Button    id:order
        Sleep    2
        ${success}=    Is Element Visible    id:order-completion
        IF    $success    BREAK
    END

Create Pdf
    [Arguments]    ${screenshotfilepath}=screenshotfilepath    ${pdffilepath}=pdffilepath
    Scroll Element Into View    id:robot-preview-image
    Capture Element Screenshot    id:robot-preview-image    ${screenshotfilepath}
    Sleep    1
    ${exists}=    Does File Exist    ${screenshotfilepath}
    IF    $exists
        ${outerhtml}=    Get Element Attribute    id:receipt    outerHTML
        Html To Pdf    ${outerhtml}    ${pdffilepath}
        ${files}=    Create List    ${screenshotfilepath}
        Add Files To Pdf    ${files}    ${pdffilepath}    ${TRUE}
        Sleep    1
        ${exists}=    Does File Exist    ${pdffilepath}
    END

Handle All Orders
    [Arguments]    ${csvfilepath}=csvfilepath    ${orderscsvurl}=orderscsvurl    ${pdffolder}=pdffolder
    ${pdffolder}=    Set Variable    ${robotroot}pdfs${/}
    ${zippath}=    Set Variable    ${robotroot}zippedpdfs.zip
    Download Orders File    ${orderscsvurl}    targetfilepath=${csvfilepath}
    ${orders}=    Read table from CSV    ${csvfilepath}    ${TRUE}
    Log    ${orders}
    ${screenshotfilepath}=    Set Variable    ${outputfolder}/testi.png
    FOR    ${order}    IN    @{orders}
        Build Robot
        ...    ${order}[Head]
        ...    bodynumber=${order}[Body]
        ...    legsnumber=${order}[Legs]
        ...    address=${order}[Address]
        ${ordernumber}=    Set Variable    ${order}[Order number]
        ${fileextension}=    Set Variable    .pdf
        ${pdffilepath}=    Set Variable    ${pdffolder}${ordernumber}${fileextension}
        Create Pdf    ${screenshotfilepath}    pdffilepath=${pdffilepath}
        Click Button    id:order-another
    END

Archive Pdfs
    ${pdffolder}=    Set Variable    ${robotroot}pdfs${/}
    ${zippath}=    Set Variable    ${robotroot}zippedpdfs.zip
    Archive Folder With Zip    ${pdffolder}    ${zippath}    ${FALSE}
    Move File    ${zippath}    ${outputfolder}zippedpdfs.zip    ${TRUE}
