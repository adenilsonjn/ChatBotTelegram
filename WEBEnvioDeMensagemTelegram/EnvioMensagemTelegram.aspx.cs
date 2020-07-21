using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Telegram.Bot.Types.Enums;

namespace WEBEnvioDeMensagemTelegram
{
    public partial class EnvioMensagemTelegram : System.Web.UI.Page
    {


        protected void Page_Load(object sender, EventArgs e)
        {

            alertaP1P2();

        }

        private void alertaP1P2()
        {
            //Planilha Teste
            //OleDbConnection ab = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Adenilson\WEBEnvioMensagemTelegram\WEBEnvioDeMensagemTelegram\WEBEnvioDeMensagemTelegram\Planilha\incidente.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;ReadOnly=False';");

            //Planilha Produção
            OleDbConnection ab = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Adenilson\WEBEnvioMensagemTelegram\WEBEnvioDeMensagemTelegram\WEBEnvioDeMensagemTelegram\Planilha\Atividades.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;ReadOnly=False';");
            //OleDbConnection ab = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\appp\Planilha\incident.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;ReadOnly=False';");
            ab.Open();

            //Select Teste
            //string select = "SELECT * FROM [Sheet1$]";

            //Select Produção
            string select = "SELECT * FROM [_Base$] WHERE incident_state IN ('Atribuído','Calendarizado','Em Resolução','Pendente Corrective Change','Pendente de Utilizador')";

            OleDbCommand comando = new OleDbCommand(select, ab);
            OleDbDataReader b = comando.ExecuteReader();

            string table = "";


            while (b.Read())
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                //Torre Distribuídos

                try
                {
                    if (b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Onsite Billing" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Hemera" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-PIM" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Iris" ||
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Easy Way- DIPJ" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Ecomex" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-E-Process" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Glorian" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Hemera" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Iris" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Magnitude" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Onsite Billing" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-OSE - On-site entrega" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-PIM" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Ponto Eletrônico" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-SGL-EDPBR" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Sistema Nexo" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-SML" || b["assignment_group"].ToString() == "EDPBR-PRO-Logica-Pacotes" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Gauss" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-GENE" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-GedWeb" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Controle de Filas" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Singular" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Risk Control" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-CHAT/SACMOBILEDESK" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-OSER" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Nota Fiscal Eletrônica" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Portal de Governança EDP" || b["assignment_group"].ToString() == "EDPBR-PRO-EVO-Accenture-OSE - On-site entrega" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Portal de Serviços" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-GIROWeb" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Sinergie" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE-RAID" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE- SAP CONCUR" || b["assignment_group"].ToString() == "EDP-PRO-Accenture SIG-GAS" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-ACCENTURE-DISTRIBUÍDOS" || b["assignment_group"].ToString() == "EDPBR-PRO-Agência Virtual" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTUR-RAID" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE PW-SATI ADM/FINANC" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE PW-SATI COMERCIAL" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ADP" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Agência Virtual" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-CAT" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-CCK" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ClientSCDE" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-CRM Enertrade" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Cubeplan" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Digitalização Assinatura de Contratos" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE-EVIDTOOL" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-EVIEWS" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Good Card" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Governanca" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE-INDQUAL" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Insighter-EDPBR" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Modeler-EDPBR" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE-ONSITEWEB" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Relogio de Ponto" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Rhevolution" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-RONDA ACESSO" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Rotas de Carreira" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-SAP HR - Historico" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE-SAP HR HISTORICO" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-SAP SNC" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Siscotrader" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Sistema de Viagem" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-SMS" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Tarifação Bancária" || 
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-URA" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE-WEBNSONAR" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Wedo (RAID)" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-WPA")
                    {
                        if (b["incident_state"].ToString() == "Atribuído" || b["incident_state"].ToString() == "Calendarizado" || b["incident_state"].ToString() == "Em Resolução" || b["incident_state"].ToString() == "Pendente Corrective Change" || b["incident_state"].ToString() == "Pendente de Utilizador")
                        {
                            if (b["priority"].ToString() == "1 - Crítico" || b["priority"].ToString() == "2 - Alto")
                            {

                                Telegram.Bot.TelegramBotClient telegramBot = new Telegram.Bot.TelegramBotClient("1101460808:AAEZB9ZBxBr0CrdVqJ1RnXM324Ow4WVOzDA");
                                telegramBot.SendTextMessageAsync(chatId: "-1001414606335", text: "Incidente: " + b["number"].ToString().TrimEnd() + "\nPrioridade: " + b["priority"].ToString().TrimEnd() + "\nStatus: " + b["incident_state"].ToString() + "\nAberto em: " + b["opened_at"].ToString() + "\nFila: " + b["assignment_group"].ToString(), parseMode: ParseMode.Html);
                                lblMensagem.Text = "Mensagem enviada com sucesso!";

                                table += "<tr>";
                                table += "<td>" + b["number"].ToString() + "</td>";
                                table += "<td>" + b["opened_at"].ToString() + "</td>";
                                table += "<td>" + b["incident_state"].ToString().TrimEnd() + "</td>";
                                table += "<td>" + b["priority"].ToString() + "</td>";
                                table += "<td>" + b["u_subcategory"].ToString() + "</td>";
                                table += "<td>" + b["assignment_group"].ToString() + "</td>";
                                table += "</tr>";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    lblMensagem.Text = "Erro ao enviar a messagem!" + ex;
                }


                //Torre Atendimento
                try
                {
                    if (b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-CRM B2C" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-CRM SALESFORCE" ||
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ECC - HR" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ECC - PM" ||
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ECC - PS" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ECC - SD" ||
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-IS-U/CSS - CRM" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-ACCENTURE-LUMUS")
                    {
                        if (b["incident_state"].ToString() == "Atribuído" || b["incident_state"].ToString() == "Calendarizado" || b["incident_state"].ToString() == "Em Resolução" || b["incident_state"].ToString() == "Pendente Corrective Change" || b["incident_state"].ToString() == "Pendente de Utilizador")
                        {
                            if (b["priority"].ToString() == "1 - Crítico" || b["priority"].ToString() == "2 - Alto")
                            {

                                Telegram.Bot.TelegramBotClient telegramBot = new Telegram.Bot.TelegramBotClient("1101460808:AAEZB9ZBxBr0CrdVqJ1RnXM324Ow4WVOzDA");
                                telegramBot.SendTextMessageAsync(chatId: "-1001380403078", text: "Incidente: " + b["number"].ToString().TrimEnd() + "\nPrioridade: " + b["priority"].ToString().TrimEnd() + "\nStatus: " + b["incident_state"].ToString() + "\nAberto em: " + b["opened_at"].ToString() + "\nFila: " + b["assignment_group"].ToString(), parseMode: ParseMode.Html);
                                lblMensagem.Text = "Mensagem enviada com sucesso!";

                                table += "<tr>";
                                table += "<td>" + b["number"].ToString() + "</td>";
                                table += "<td>" + b["opened_at"].ToString() + "</td>";
                                table += "<td>" + b["incident_state"].ToString().TrimEnd() + "</td>";
                                table += "<td>" + b["priority"].ToString() + "</td>";
                                table += "<td>" + b["u_subcategory"].ToString() + "</td>";
                                table += "<td>" + b["assignment_group"].ToString() + "</td>";
                                table += "</tr>";

                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    lblMensagem.Text = "Erro ao enviar a messagem!" + ex;
                }



                //Torre Arrecadação, Cobrança e Faturamento

                try
                {
                    if (b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-IS-U/CCS - BILLING" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-IS-U/CCS - DM" ||
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-IS-U/CCS - FICA" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ECC - FI" ||
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ECC - CO")
                    {
                        if (b["incident_state"].ToString() == "Atribuído" || b["incident_state"].ToString() == "Calendarizado" || b["incident_state"].ToString() == "Em Resolução" || b["incident_state"].ToString() == "Pendente Corrective Change" || b["incident_state"].ToString() == "Pendente de Utilizador")
                        {
                            if (b["priority"].ToString() == "1 - Crítico" || b["priority"].ToString() == "2 - Alto")
                            {
                                Telegram.Bot.TelegramBotClient telegramBot = new Telegram.Bot.TelegramBotClient("1101460808:AAEZB9ZBxBr0CrdVqJ1RnXM324Ow4WVOzDA");
                                telegramBot.SendTextMessageAsync(chatId: "-1001381132694", text: "Incidente: " + b["number"].ToString().TrimEnd() + "\nPrioridade: " + b["priority"].ToString().TrimEnd() + "\nStatus: " + b["incident_state"].ToString() + "\nAberto em: " + b["opened_at"].ToString() + "\nFila: " + b["assignment_group"].ToString(), parseMode: ParseMode.Html);
                                lblMensagem.Text = "Mensagem enviada com sucesso!";

                                table += "<tr>";
                                table += "<td>" + b["number"].ToString() + "</td>";
                                table += "<td>" + b["opened_at"].ToString() + "</td>";
                                table += "<td>" + b["incident_state"].ToString().TrimEnd() + "</td>";
                                table += "<td>" + b["priority"].ToString() + "</td>";
                                table += "<td>" + b["u_subcategory"].ToString() + "</td>";
                                table += "<td>" + b["assignment_group"].ToString() + "</td>";
                                table += "</tr>";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    lblMensagem.Text = "Erro ao enviar a messagem!" + ex;
                }


                //Torre Serviço de Campo

                try
                {
                    if (b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-IS-U/CCS - WM" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-IS-U/CCS - DM" ||
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ECC - MM" || b["assignment_group"].ToString() == "EDPBR-PROJ-ACCENTURE-WMS")
                    {
                        if (b["incident_state"].ToString() == "Atribuído" || b["incident_state"].ToString() == "Calendarizado" || b["incident_state"].ToString() == "Em Resolução" || b["incident_state"].ToString() == "Pendente Corrective Change" || b["incident_state"].ToString() == "Pendente de Utilizador")
                        {
                            if (b["priority"].ToString() == "1 - Crítico" || b["priority"].ToString() == "2 - Alto")
                            {
                                Telegram.Bot.TelegramBotClient telegramBot = new Telegram.Bot.TelegramBotClient("1101460808:AAEZB9ZBxBr0CrdVqJ1RnXM324Ow4WVOzDA");
                                telegramBot.SendTextMessageAsync(chatId: "-1001296061125", text: "Incidente: " + b["number"].ToString().TrimEnd() + "\nPrioridade: " + b["priority"].ToString().TrimEnd() + "\nStatus: " + b["incident_state"].ToString() + "\nAberto em: " + b["opened_at"].ToString() + "\nFila: " + b["assignment_group"].ToString(), parseMode: ParseMode.Html);
                                lblMensagem.Text = "Mensagem enviada com sucesso!";

                                table += "<tr>";
                                table += "<td>" + b["number"].ToString() + "</td>";
                                table += "<td>" + b["opened_at"].ToString() + "</td>";
                                table += "<td>" + b["incident_state"].ToString().TrimEnd() + "</td>";
                                table += "<td>" + b["priority"].ToString() + "</td>";
                                table += "<td>" + b["u_subcategory"].ToString() + "</td>";
                                table += "<td>" + b["assignment_group"].ToString() + "</td>";
                                table += "</tr>";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    lblMensagem.Text = "Erro ao enviar a messagem!" + ex;
                }


                //Torre Analytics
                try
                {
                    if (b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-SAP BPC" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-SAP BW")
                    {
                        if (b["incident_state"].ToString() == "Atribuído" || b["incident_state"].ToString() == "Calendarizado" || b["incident_state"].ToString() == "Em Resolução" || b["incident_state"].ToString() == "Pendente Corrective Change" || b["incident_state"].ToString() == "Pendente de Utilizador")
                        {
                            if (b["priority"].ToString() == "1 - Crítico" || b["priority"].ToString() == "2 - Alto")
                            {
                                Telegram.Bot.TelegramBotClient telegramBot = new Telegram.Bot.TelegramBotClient("1101460808:AAEZB9ZBxBr0CrdVqJ1RnXM324Ow4WVOzDA");
                                telegramBot.SendTextMessageAsync(chatId: "-1001121289892", text: "Incidente: " + b["number"].ToString().TrimEnd() + "\nPrioridade: " + b["priority"].ToString().TrimEnd() + "\nStatus: " + b["incident_state"].ToString() + "\nAberto em: " + b["opened_at"].ToString() + "\nFila: " + b["assignment_group"].ToString(), parseMode: ParseMode.Html);
                                lblMensagem.Text = "Mensagem enviada com sucesso!";

                                table += "<tr>";
                                table += "<td>" + b["number"].ToString() + "</td>";
                                table += "<td>" + b["opened_at"].ToString() + "</td>";
                                table += "<td>" + b["incident_state"].ToString().TrimEnd() + "</td>";
                                table += "<td>" + b["priority"].ToString() + "</td>";
                                table += "<td>" + b["u_subcategory"].ToString() + "</td>";
                                table += "<td>" + b["assignment_group"].ToString() + "</td>";
                                table += "</tr>";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    lblMensagem.Text = "Erro ao enviar a messagem!" + ex;
                }

                //Torre Técnica
                try
                {
                    if (b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-ABAP" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-SAP PI" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-SAP GRC" ||
                        b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Charm" || b["assignment_group"].ToString() == "EDPBR-PRO-CORR-Accenture-Solution Manager")
                    {
                        if (b["incident_state"].ToString() == "Atribuído" || b["incident_state"].ToString() == "Calendarizado" || b["incident_state"].ToString() == "Em Resolução" || b["incident_state"].ToString() == "Pendente Corrective Change" || b["incident_state"].ToString() == "Pendente de Utilizador")
                        {
                            if (b["priority"].ToString() == "1 - Crítico" || b["priority"].ToString() == "2 - Alto")
                            {
                                Telegram.Bot.TelegramBotClient telegramBot = new Telegram.Bot.TelegramBotClient("1101460808:AAEZB9ZBxBr0CrdVqJ1RnXM324Ow4WVOzDA");
                                telegramBot.SendTextMessageAsync(chatId: "-1001235945109", text: "Incidente: " + b["number"].ToString().TrimEnd() + "\nPrioridade: " + b["priority"].ToString().TrimEnd() + "\nStatus: " + b["incident_state"].ToString() + "\nAberto em: " + b["opened_at"].ToString() + "\nFila: " + b["assignment_group"].ToString(), parseMode: ParseMode.Html);
                                lblMensagem.Text = "Mensagem enviada com sucesso!";

                                table += "<tr>";
                                table += "<td>" + b["number"].ToString() + "</td>";
                                table += "<td>" + b["opened_at"].ToString() + "</td>";
                                table += "<td>" + b["incident_state"].ToString().TrimEnd() + "</td>";
                                table += "<td>" + b["priority"].ToString() + "</td>";
                                table += "<td>" + b["u_subcategory"].ToString() + "</td>";
                                table += "<td>" + b["assignment_group"].ToString() + "</td>";
                                table += "</tr>";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    lblMensagem.Text = "Erro ao enviar a messagem!" + ex;
                }

            }

            PlTable.Controls.Add(new LiteralControl(table));

            b.Close();
            ab.Close();
        }
    }
}