*** Settings ***
Documentation   Template robot main suite.
Library         SeleniumLibrary
Library         Collections
Library         libraries/ExampleHelper.py
Resource        keywords/keywords.robot
Library    RPA.PDF
Library    String
Library    RPA.Browser.Selenium
Library    RPA.Desktop
Library    libraries/o365.py
Library    OperatingSystem

*** Keywords ***
Leer PDF Y Extrae DNI
    [Arguments]    ${fichero}
    ${pdf}=    Get Text From Pdf    ${CURDIR}${/}${fichero}
    Log    ${pdf}
    ${texto}=    Get Dictionary Values    ${pdf}
    Log    ${texto}[0]
    ${DNI}=    Get Regexp Matches    ${texto}[0]    [0-9]{8}.
    Log    ${DNI}
    [Return]    ${DNI}
Enviar por vidsigner
    [Arguments]    ${DNI}    ${Email_Vid}    ${Pass_Vid}
    Open Chrome Browser    https://limcamar.send2sign.net    maximized=${True}
    RPA.Browser.Selenium.Input Text    id=email    ${Email_Vid}
    RPA.Browser.Selenium.Input Text    id=password    ${Pass_Vid}
    RPA.Browser.Selenium.Wait Until Element Is Visible  xpath=//input[@value='Iniciar sesión']  130
    RPA.Browser.Selenium.Click Element    xpath=//input[@value='Iniciar sesión']
    RPA.Browser.Selenium.Wait Until Element Is Visible    css=.fa-share-square    55
    RPA.Browser.Selenium.Click Element    css=.fa-share-square
    RPA.Browser.Selenium.Wait Until Element Is Visible    css=.modal-content:nth-child(1) .btn    55
    RPA.Browser.Selenium.Click Element   css=.modal-content:nth-child(1) .btn
    RPA.Browser.Selenium.Choose File    name=files[]    ${CURDIR}${/}Certificado_de_aprovechamiento-ESTELA_MARIA_RAINERO_GOMEZ.pdf
    RPA.Browser.Selenium.Wait Until Element Is Visible    name=editShortDesc    55
    RPA.Browser.Selenium.Input Text    name=editShortDesc    ${DNI}
    RPA.Browser.Selenium.Input Text    xpath=/html/body/div[1]/div[3]/div[3]/div/div/div/div[1]/div/div[4]/div/section[1]/form/div[2]/div[4]/div[1]/textarea       CERTIFICADO FORMACION
    RPA.Browser.Selenium.Select From List By Index     xpath=//*[contains(@id, 'editTemplate')]   1
    RPA.Browser.Selenium.Wait Until Element Is Visible    xpath=//*[contains(@id, 'btn-plus-sig')]    55
    RPA.Browser.Selenium.Click Element    xpath=//*[contains(@id, 'btn-plus-sig')]
    #RPA.Browser.Selenium.Click Element   css=.fa-plus-circle
    RPA.Browser.Selenium.Click Element    id=showFiltersButton
    RPA.Browser.Selenium.Input Text    xpath=/html/body/div[1]/div[3]/div[3]/div/div/div/div[1]/div/div[7]/div/div/div[2]/div[1]/div/div[1]/div/table/thead/tr[2]/th[3]/span/span/span[1]/input   PABLO SAURA
    RPA.Browser.Selenium.Press Keys    xpath=/html/body/div[1]/div[3]/div[3]/div/div/div/div[1]/div/div[7]/div/div/div[2]/div[1]/div/div[1]/div/table/thead/tr[2]/th[3]/span/span/span[1]/input       ENTER
    RPA.Browser.Selenium.Wait Until Element Is Visible    css=td:nth-child(3)    55
    RPA.Browser.Selenium.Click Element    css=td:nth-child(3)
    RPA.Browser.Selenium.Wait Until Element Is Visible    xpath=//*[contains(@id, 'btn-savePos')]    55
    RPA.Browser.Selenium.Click Element    xpath=//*[contains(@id, 'btn-savePos')]
    RPA.Browser.Selenium.Click Button    id=btn-save-send    
  


*** Tasks ***
Principal
    ${email}    Get Environment Variable        email
    ${tennant}    Get Environment Variable        tennant
    ${user}    Get Environment Variable        user
    ${key}    Get Environment Variable        key
    ${Email_Vid}    Get Environment Variable        Email_Vid
    ${Pass_Vid}    Get Environment Variable        Pass_Vid
    Log     ${email}
    ${documento}=    Credenciales O365     ${CURDIR}    ${user}    ${key}    ${tennant}    ${email}
    ${DNI}=    Leer PDF Y Extrae DNI    Certificado_de_aprovechamiento-ESTELA_MARIA_RAINERO_GOMEZ.pdf
    Enviar por vidsigner    ${DNI}   ${Email_Vid}    ${Pass_Vid}    


    
