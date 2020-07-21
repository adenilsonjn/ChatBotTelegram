<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="EnvioMensagemTelegram.aspx.cs" Inherits="WEBEnvioDeMensagemTelegram.EnvioMensagemTelegram" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>

    <script type="text/javascript">
        function Atualizar() {
            window.location.reload();
        }
    </script>

</head>
<body onload="setInterval('Atualizar()', 60000)">
    <form id="form1" runat="server">
        <div>
            <asp:Label ID="lblMensagem" runat="server"></asp:Label>
        </div>
        <div class="card-body">
            <table id="example1" class="table table-bordered table-striped">
                <thead>
                    <tr>
                        <th>Numero</th>
                        <th>Aberto</th>
                        <th>Prioridade</th>
                        <th>Subcategoria</th>
                        <th>Fila</th>
                    </tr>
                </thead>
                <tbody>
                    <asp:Panel ID="PlTable" runat="server"></asp:Panel>
                </tbody>
                <tfoot>
                    <tr>
                        <th>Numero</th>
                        <th>Aberto</th>
                        <th>Prioridade</th>
                        <th>Subcategoria</th>
                        <th>Fila</th>
                    </tr>
                </tfoot>
            </table>
        </div>

    </form>
</body>
</html>
