Attribute VB_Name = "modChrome"
Sub MudarOpcaoParaDownloadPDF()
''
'' Muda a opção do Chrome para fazer download, em vez de abrir no navegador.
''

    Dim chrome As Selenium.ChromeDriver
    Set chrome = New Selenium.ChromeDriver
    chrome.Start "chrome"
    
    ' Abre a tela de configurações
    chrome.Get "chrome://settings/content/pdfDocuments"
    
    If chrome.FindElementById("control").IsSelected = False Then
        chrome.FindElementById("control").Click
    End If
    
    chrome.Close ' Close ou quit?
    
End Sub


Sub SetarPreferenciasChrome(strDiretorio)
''
'' Configura as preferências do Chrome para não perguntar se deseja salvar o download e para salvar na pasta strDiretorio.
''

    Dim chrome As Selenium.ChromeDriver
    Set chrome = New Selenium.ChromeDriver
    chrome.Start "chrome"
    
    ' Desabilitar popup perguntando onde salvar
    chrome.SetPreference "profile.default_content_settings.popups", 0
    chrome.SetPreference "download.prompt_for_download", "false"
    chrome.SetPreference "download.directory_upgrade", True
    'driver.SetPreference "safebrowsing.enabled", True
    chrome.SetPreference "plugins.plugins_disabled", Array("Chrome PDF Viewer")
    
    ' Configurar o diretório para download
    chrome.SetPreference "download.default_directory", strDiretorio
    
End Sub
