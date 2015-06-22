Option Explicit

Function testjs()
    Dim js As New cJavaScript, encrypted As String, decrypted
    
    With js
        ' not really necessary first time in
        .clear

        ' add libraries - I want to try a few different kinds of encryption
        .addUrl "http://crypto-js.googlecode.com/svn/tags/3.0.2/build/rollups/aes.js"
        .addUrl "http://crypto-js.googlecode.com/svn/tags/3.0.2/build/rollups/tripledes.js"
        .addUrl "http://crypto-js.googlecode.com/svn/tags/3.0.2/build/rollups/rabbit.js"
        
        ' add my code
        .addCode _
            "function encrypt(msg, pass, method) {" & _
            "    return CryptoJS[method].encrypt(msg, pass).toString();" & _
            "}" & _
            "function decrypt(encryptedMessage, pass,method) {" & _
            "    return CryptoJS[method].decrypt(encryptedMessage, pass, method).toString(CryptoJS.enc.Utf8);" & _
            "}"

        'various encryptions for fun
        encrypted = .compile.run("encrypt", "a message from aes", "my passphrase", "AES")
        decrypted = .compile.run("decrypt", encrypted, "my passphrase", "AES")
        Debug.Print encrypted, decrypted
        
        encrypted = .compile.run("encrypt", "a message from tripledes", "my passphrase", "TripleDES")
        decrypted = .compile.run("decrypt", encrypted, "my passphrase", "TripleDES")
        Debug.Print encrypted, decrypted
        
        encrypted = .compile.run("encrypt", "a message from rabbit", "my passphrase", "Rabbit")
        decrypted = .compile.run("decrypt", encrypted, "my passphrase", "Rabbit")
        Debug.Print encrypted, decrypted
    End With

End Function
