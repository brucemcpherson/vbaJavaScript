Option Explicit
'<script src="https://cdn.polyfill.io/v1/polyfill.min.js"></script>
Private Function testGas()
    Dim js As New cJavaScript, start As Double, numberOfTests As Long, result As Variant
    numberOfTests = 10000
    
    With js
        ' not really necessary first time in
        .clear

        ' here's a couple of polyfills to bring it more or less up to apps-script levels
        .addUrl "https://cdnjs.cloudflare.com/ajax/libs/json2/20150503/json2.min.js"
        .addUrl "https://cdnjs.cloudflare.com/ajax/libs/es5-shim/4.1.7/es5-shim.min.js"
        
        ' get my code from apps script
        .addAppsScript "https://script.google.com/macros/s/AKfycbzVhdyNg3-9jBu6KSLkYIwN48vuXCp6moOLQzQa7eXar7HdWe8/exec?manifests=color"

        ' my code
        .addCode ("function compareColors (rgb1, rgb2) { " & _
                " return theColorProp(rgb1).compareColorProps(theColorProp(rgb2).getProperties()) ; " & _
            "}" & _
            "function compareColorTest (numberOfTests) {" & _
                "for (var i = 0 , t = 0 ; i < numberOfTests ; i++ ) { " & _
                "   t += compareColors ( Math.round(Math.random() * VBCOLORS.vbWhite) , Math.round(Math.random() * VBCOLORS.vbWhite) ); " & _
                "}" & _
                " return 'average color distance:' + t/i;" & _
            "}" & _
            "function theColorProp (rgb1) { " & _
                " return new ColorMath(rgb1) ; " & _
            "}" & _
            "function theColorPropStringified (rgb1) { " & _
                " return JSON.stringify(theColorProp(rgb1).getProperties()) ; " & _
            "}")
        
        'a stringified color properties
        Debug.Print .compile.run("theColorPropStringified", vbBlue)
      
        'compare a couple of colors
        Debug.Print .compile.run("compareColors", vbBlue, vbRed)
        
        ' do a performance test
        start = tinyTime
        result = .compile.run("compareColorTest", numberOfTests)
        Debug.Print "time to complete in JS " & (tinyTime - start)
        Debug.Print result
      
    End With

End Function
Private Function comparePerformance()
    Dim js As New cJavaScript, result As Variant, start As Double, numberOfTests As Long, i As Long, t As Double
    numberOfTests = 20000
    
    With js
        ' not really necessary first time in
        .clear

        ' its old enough not to have JSON.parse/.stringify - so we'll polyfill that
        .addUrl "https://cdnjs.cloudflare.com/ajax/libs/json2/20150503/json2.min.js"
        
        ' get my apps script code from git hub
        .addUrl "https://raw.githubusercontent.com/brucemcpherson/ColorArranger/master/scripts/ColorMath.js.html"
        
        ' apps script html/js files often have script tags embedded
        .removeScriptTags
        
        ' my code
        .addCode ("function compareColors (rgb1, rgb2) { " & _
                " return theColorProp(rgb1).compareColorProps(theColorProp(rgb2).getProperties()) ; " & _
            "}" & _
            "function compareColorTest (numberOfTests) {" & _
                "for (var i = 0 , t = 0 ; i < numberOfTests ; i++ ) { " & _
                "   t += compareColors ( Math.round(Math.random() * VBCOLORS.vbWhite) , Math.round(Math.random() * VBCOLORS.vbWhite) ); " & _
                "}" & _
                " return 'average color distance:' + t/i;" & _
            "}" & _
            "function theColorProp (rgb1) { " & _
                " return new ColorMath(rgb1) ; " & _
            "}" & _
            "function theColorPropStringified (rgb1) { " & _
                " return JSON.stringify(theColorProp(rgb1).getProperties()) ; " & _
            "}")
        
        'a stringified color properties
        Debug.Print .compile.run("theColorPropStringified", vbBlue)
      
        'compare a couple of colors
        Debug.Print .compile.run("compareColors", vbBlue, vbRed)
        
        ' do a performance test
        start = tinyTime
        result = .compile.run("compareColorTest", numberOfTests)
        Debug.Print "time to complete in JS " & (tinyTime - start)
        Debug.Print result
    End With
    
    ' now lets do the same thing in native VBA - can't easily stringify a custom type so I'll just do one prop of it
    Debug.Print makeColorProps(vbBlue).htmlHex
    
    'compare a couple of colors
    Debug.Print compareColors(vbBlue, vbRed)
    
    ' compare loads of colors and time it
    start = tinyTime
    t = 0
    For i = 1 To numberOfTests
        t = t + compareColors(CLng(Round(Rnd() * vbWhite)), CLng(Round(Rnd() * vbWhite)))
    Next i

    Debug.Print "time to complete in VBA " & (tinyTime - start)
    Debug.Print "Average Color distance: " & (t / numberOfTests)
    
End Function
Private Function testGit()
    Dim js As New cJavaScript, result As String
    
    With js
        ' not really necessary first time in
        .clear

        ' its old enough not to have JSON.parse/.stringify - so we'll polyfill that
        .addUrl "https://cdnjs.cloudflare.com/ajax/libs/json2/20150503/json2.min.js"
        
        ' get my apps script code from git hub
        .addUrl "https://raw.githubusercontent.com/brucemcpherson/ColorArranger/master/scripts/ColorMath.js.html"
        
        ' apps script html/js files often have script tags embedded
        .removeScriptTags
        
        ' my code
        .addCode "function getColorProperties (hex) { " & _
            " return JSON.stringify(VBCOLORS); " & _
         "}"
        
        'try these out
        result = .compile.run("myTest")
      
        Debug.Print result
    End With

End Function

Private Function testJs()
    Dim js As New cJavaScript, encrypted As String, decrypted As String
    
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
        Debug.Print decrypted, encrypted
        
        encrypted = .compile.run("encrypt", "a message from tripledes", "my passphrase", "TripleDES")
        decrypted = .compile.run("decrypt", encrypted, "my passphrase", "TripleDES")
        Debug.Print decrypted, encrypted
        
        encrypted = .compile.run("encrypt", "a message from rabbit", "my passphrase", "Rabbit")
        decrypted = .compile.run("decrypt", encrypted, "my passphrase", "Rabbit")
        Debug.Print decrypted, encrypted
    End With

End Function

Private Function jsvsvbaTests()
    Dim js As New cJavaScript, i As Long, vbaResult As Double, _
        javaScriptResult As Double, rgb1 As Long, rgb2 As Long, _
        numberOfTests As Long
    
    numberOfTests = 10000
    With js
        ' not really necessary first time in
        .clear

        ' here's a couple of polyfills to bring it more or less up to apps-script levels
        .addUrl "https://cdnjs.cloudflare.com/ajax/libs/json2/20150503/json2.min.js"
        .addUrl "https://cdnjs.cloudflare.com/ajax/libs/es5-shim/4.1.7/es5-shim.min.js"
        
        ' get my code directly from apps script
        .addAppsScript "https://script.google.com/macros/s/AKfycbzVhdyNg3-9jBu6KSLkYIwN48vuXCp6moOLQzQa7eXar7HdWe8/exec?manifests=color"
        
        ' add my unit test code
        .addCode ("function compareColors (rgb1, rgb2) { " & _
                        " return new ColorMath(rgb1).compareColorProps(new ColorMath(rgb2).getProperties()) ; " & _
                    "}")
            
        
        'well run lots of random tests and check they are the same
        For i = 1 To numberOfTests
            
            ' get some test data
            rgb1 = CLng(Round(Rnd() * vbWhite))
            rgb2 = CLng(Round(Rnd() * vbWhite))
            
            ' run it in VBA
            vbaResult = compareColors(rgb1, rgb2)
            
            ' run it in javaScript
            javaScriptResult = .compile.run("compareColors", rgb1, rgb2)
            
            ' should be the same result
            ' but need to round a bit as you cant compare doubles for exact equality
            ' so we'll compare to 8 decimal places
            Debug.Assert Round(vbaResult, 8) = Round(javaScriptResult, 8)
        Next i

     

        
    End With
End Function