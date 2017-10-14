<?
// aqui comeca o script
//pega as variaveis por POST
$nome      = $_POST["nome"];
$email   = $_POST["email"];
$assunto   = $_POST["assunto"];
$mensagem  = $_POST["mensagem"];

global $email; //funcao para validar a variavel $email no script todo

$data      = date("d/m/y");                     //funcao para pegar a data de envio do e-mail
$ip        = $_SERVER['REMOTE_ADDR'];           //funcao para pegar o ip do usu&#65533;rio
$navegador = $_SERVER['HTTP_USER_AGENT'];       //funcao para pegar o navegador do visitante
$hora      = date("H:i");                       //para pegar a hora com a funcao date

// Verifica se o campo nome ta preenchido
if (empty($nome)){
// HTML que aparecera o ERRO
echo "<html><head><title>Ocorreu Um ERRO!!!</title></head>";
echo "<body bgcolor=\"#ffffff\">";
echo "<br><br><br>";
echo "<center>&Eacute; Necess&aacute;rio o Preenchimento do <b>Nome</b></center>";
echo "<br><br><center><a href=\"javascript:history.back(1)\">Voltar</a></center>";
echo "</body></html>";
}
// Verifica o campo E-mail ta preenchido
elseif (empty($email)){
// HTML que aparecera o ERRO
echo "<html><head><title>Ocorreu Um ERRO!!!</title></head>";
echo "<body bgcolor=\"#ffffff\">";
echo "<br><br><br>";
echo "<center>O E-mail n&atilde;o foi <b>Digitado</b></center>";
echo "<br><br><center><a href=\"javascript:history.back(1)\">Voltar</a></center>";
echo "</body></html>";
}
// Verifica se o E-mail Contem @
elseif (!(strpos($email,"@")) OR strpos($email,"@") !=strrpos($email,"@")) {
// HTML que aparecera o ERRO
echo "<html><head><title>Ocorreu Um ERRO!!!</title></head>";
echo "<body bgcolor=\"#ffffff\">";
echo "<br><br><br>";
echo "<center>O E-mail Digitado <b>N&atilde;o</b> &eacute; <b>v&aacute;lido</b></center>";
echo "<br><br><center><a href=\"javascript:history.back(1)\">Voltar</a></center>";
echo "</body></html>";
}
// Verifoca Se o E-mail Contem .
elseif (!(strpos($email,".")) OR strpos($email,".") !=strrpos($email,".")) {
// HTML que aparecera o ERRO
echo "<html><head><title>Ocorreu Um ERRO!!!</title></head>";
echo "<body bgcolor=\"#ffffff\">";
echo "<br><br><br>";
echo "<center>O E-mail Digitado <b>N&atilde;o</b> &eacute; <b>v&aacute;lido</b></center>";
echo "<br><br><center><a href=\"javascript:history.back(1)\">Voltar</a></center>";
echo "</body></html>";
}
// Verifica se o campo Assunto ta preenchido
elseif (empty($assunto)){
// HTML que aparecera o ERRO
echo "<html><head><title>Ocorreu Um ERRO!!!</title></head>";
echo "<body bgcolor=\"#ffffff\">";
echo "<br><br><br>";
echo "<center>Voc&ecirc; <b>N&atilde;o</b> Digitou Um <b>Assunto</b></center>";
echo "<br><br><center><a href=\"javascript:history.back(1)\">Voltar</a></center>";
echo "</body></html>";
}
// Verifica se o campo Mensagem ta preenchido
elseif (empty($mensagem)){
// HTML que aparecera o ERRO
echo "<html><head><title>Ocorreu Um ERRO!!!</title></head>";
echo "<body bgcolor=\"#ffffff\">";
echo "<br><br><br>";
echo "<center>Voc&ecirc; <b>N&atilde;o</b> Digitou Uma <b>Mensagem</b></center>";
echo "<br><br><center><a href=\"javascript:history.back(1)\">Voltar</a></center>";
echo "</body></html>";
}
else{
echo "<p align=center>$nome, sua mensagem foi enviada com sucesso!</p>";
echo "<p align=center>Estaremos retornando em breve.</p>";
echo "<html><head><title>Mensagem Enviada!!!</title></head>";
echo "<body bgcolor=\"#ffffff\">";
echo "<br><br><br>";
echo "<br><br><center><input type=\"button\" value=\"OK\" onclick=\"location.href= 'http://www.assistenciabh.com.br/informatica'; \"></center>";
echo "</body></html>";

//aqui envia o e-mail para voce
mail ("paulo@assistenciabh.com.br",                       //email aonde o php vai enviar os dados do form
      "$assunto",
      "Nome: $nome\nData: $data\nIp: $ip\nNavegador: $navegador\nHora: $hora\nE-mail: $email\n\nMensagem: $mensagem",
      "From: $email"
     );

//aqui sao as configuracoes para enviar o e-mail para o visitante
$site   = "paulo@assistenciabh.com.b";                    //o e-mail que aparecera na caixa postal do visitante
$titulo = "$assunto";                  //titulo da mensagem enviada para o visitante
$msg    = "$nome, obrigado por entrar em contato conosco, em breve entraremos em contato";

//aqui envia o e-mail de auto-resposta para o visitante
mail("$email",
     "$titulo",
     "$msg",
     "From: $site"
    );
}
?>