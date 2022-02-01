<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="CARGAR_EXCEL.WebForm1" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" ></script>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" />
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"></script>
    <style>
        .mitabla {
            width :100%
        }
    </style>
</head>
<body>
   
    <form id="form1" runat="server">
         <div class="container-fluid mt-4">
                 <div class="card">
                  <div class="card-header">
                    Cargar Archivo
                  </div>
                  <div class="card-body">
                    <div class="row">
                        
                        <div class="col-sm-12">
                            <div class="form-row">
                                <div class="form-group col-sm-10">
                                  <label for="txtName">Numero de orden</label>
                                    <asp:TextBox ID="txtName" CssClass="form-control" runat="server" Width="140" />
                                    
                                </div>
                                <div class="form-group col-sm-10">
                                  <label for="FileUpload1">Archivo</label>
                                    <asp:FileUpload ID="FileUpload1" CssClass="form-control-file" runat="server" required="true"/>
                                    
                                </div>
                                <div class="form-group col-sm-2">
                                  <asp:Button ID="Button1" runat="server" Text="Cargar" CssClass="btn btn-block btn-success mt-4" OnClick="Button1_Click" />
                                </div>
                            </div>
                        </div>

                    </div>
                      <hr />
                   <div class="row">
                       <div class="col-sm-12">
                           <div class="card">
                          <div class="card-header p-1">
                            Registros Cargados
                          </div>
                          <div class="card-body">
                              <asp:GridView ID="GridView1" runat="server" CssClass="table table-bordered mitabla table-hover"></asp:GridView>
                          </div>
                        </div>
                     </div>
                       

                   </div>
                  </div>
                </div>
        </div>

        
        <%--<div>
            
            <br />
            <br />
           
        </div>
        <br />
        <div>
            <asp:Label ID="lblrespuesta" runat="server"></asp:Label>
        </div>--%>
        
    </form>
    
</body>
</html>
