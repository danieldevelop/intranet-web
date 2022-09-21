dim opcion

set sitio = CreateObject("Wscript.Shell")

opcion = inputbox("MENU DE OPCIONES" & vbCrLf & vbCrLf &_
    "1. Google" & vbCrLf &_
    "2. Historia de la informatica" & vbCrLf &_
    "3. Mi portafolio" & vbCrLf &_
    "4. Youtube." & vbCrLf &_
    "0. Salir", "Intranet Webs Sites")
opcion = CInt(opcion)


select case opcion
    case 1
        sitio.Run "https://www.google.com/"
    case 2
        sitio.Run "https://es.wikipedia.org/wiki/Inform%C3%A1tica"
    case 3
        sitio.Run "https://danieldevelop.github.io/"
    case 4
        sitio.Run "https://www.youtube.com/"
    case 0
        WScript.echo "Gracias pos su visita .)"
    case else
        WScript.echo "Opcion no valida"
end select