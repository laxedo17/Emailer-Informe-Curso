using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace Emailer_Informe_Curso.Trabajadores
{
    class EnviaEmailerInformeDetalleMatriculas
    {
        public void Enviar(string nombreFichero)
        {
            SmtpClient cliente = new SmtpClient("smtp-mail.outlook.com"); //un servidor SMTP e unha maquina que se vai a encargar de todo o proceso de envio de emails, neste caso imos usar o servidor smtp de Outlook
            cliente.Port = 587;//usamos o puerto tipico para enviar emails
            cliente.DeliveryMethod = SmtpDeliveryMethod.Network;//usando o noso equipo vai a utilizar a rede para enviar o email
            NetworkCredential credencialesRed = new NetworkCredential("78@hotmail.com", "inventado");
            cliente.EnableSsl = true;//Secure Socket Layer, isto asegura que a conexion entre o noso cliente e o servidor sera segura. Conexion segura cliente-servidor
            cliente.Credentials = credencialesRed;

            MailMessage mensaje = new MailMessage("78@hotmail.com", "17@hotmail.com"); //A direccion de correo inicial envia o email a segunda direccion email e a que recibe o email
            mensaje.Subject = "Detalles de Informe Matriculas";
            mensaje.IsBodyHtml = true;//permitenos crear un email moito mais vistoso usando etiquetas html no Body(corpo) do mensaje
            mensaje.Body = "Hola,<br><br>Adxunto a este email podes atopar o informe de detalles.<br><br>Este mensaxe autodestruirase cando digas 'somos tontos'.<br><br>Saudos e para a cama.";

            Attachment adjunto = new Attachment(nombreFichero);
            mensaje.Attachments.Add(adjunto);

            cliente.Send(mensaje);
        }
    }
}
