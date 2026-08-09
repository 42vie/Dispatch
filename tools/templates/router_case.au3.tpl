            Case "{{NAME_UPPER}}_LIST"
                _SendHttpResponse($iSocket, 200, "application/json", _{{NAME_UPPER}}_ListJSON())
                TCPCloseSocket($iSocket)
                Return
            Case "{{NAME_UPPER}}_ADD"
                _SendHttpResponse($iSocket, 200, "application/json", _{{NAME_UPPER}}_AddJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
            Case "{{NAME_UPPER}}_DELETE"
                _SendHttpResponse($iSocket, 200, "application/json", _{{NAME_UPPER}}_DeleteJSON($sBody))
                TCPCloseSocket($iSocket)
                Return
