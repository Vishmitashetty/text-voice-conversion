  Dim msg, sapi

                       msg=InputBox("Enter your text for conversion– ","By net info")
                       Set sapi=CreateObject("sapi.spvoice")
                       sapi.Speak msg