<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UploadExcel.aspx.cs" Inherits="XSyncExcelUploader.UploadExcel" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <title>XSync Excel File Upload</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css">
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"></script>
    <link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/dataTables.bootstrap4.min.css" />
    <script src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js" type="text/javascript"></script>
    <script src="https://cdn.datatables.net/1.10.16/js/dataTables.bootstrap4.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            //$("#GridView1").prepend($("<thead></thead>").append($(this).find("tr:first"))).dataTable();
        });
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div id="mastercontainer" style="background:url(https://thalappakatti.com/wp-content/themes/thalappakatti/images/banner2.png);height:100vh;">
            <div class="container py-3" style="padding-top:4% !important">
            <div style="margin-left:39%"><img  src="https://thalappakatti.com/wp-content/uploads/2018/01/thalappakatti-logo-anim.gif" rel="logo" alt="Dindigul Thalappakatti Restaurant"></div><br />
            <h2 style="color:#ffbd23" class="text-center text-uppercase">Thalappakatti Excel Uploader</h2>
            <div class="card">
                <div style="background-color:#690514 !important" class="card-header bg-primary text-uppercase text-white">
                    <h5 style="color:#ffbd23">Select Your Excel File</h5>
                </div>
                <div class="card-body">
                    <button style="margin-bottom:10px;background-color:#690514 !important;color:#ffbd23" type="button" class="btn btn-primary" data-toggle="modal" data-target="#myModal">
                        <i class="fa fa-plus-circle"></i> Add Excel
                    </button>
                    <br />
                    <asp:Label ID="lblMessage" runat="server"></asp:Label>
                    <div class="modal fade" id="myModal">
                        <div class="modal-dialog">
                            <div class="modal-content" style="width:90%">
                                <div class="modal-header">
                                    <h4 class="modal-title">Import Excel File</h4>
                                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                                </div>
                                <div class="modal-body">
                                    <div class="row">
                                        <div class="col-md-12">
                                            <div class="form-group">
                                                <!--<label>Click Browse below</label>-->
                                                <div class="input-group">
                                                    <div class="custom-file">
                                                        <asp:FileUpload Visible="true" ID="FileUpload1" runat="server" />
                                                        <!--<label class="custom-file-label"></label>-->
                                                    </div>
                                                    <label id="filename"></label>
                                                    <div class="input-group-append">
                                                        <asp:Button ID="btnUpload" runat="server" BackColor="#690514" ForeColor="#ffbd23" CssClass="btn btn-outline-primary" Text="Upload" OnClick="btnUpload_Click" />
                                                    </div>
                                                </div>
                                                
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div class="modal-footer">
                                    <button type="button" style="background-color:#690514 !important;color:#ffbd23 !important" class="btn btn-danger" data-dismiss="modal">Close</button>
                                </div>
                            </div>
                        </div>
                    </div>
                   <!-- <asp:GridView ID="GridView1" HeaderStyle-CssClass="bg-primary text-white" ShowHeaderWhenEmpty="true" runat="server" AutoGenerateColumns="false" CssClass="table table-bordered``">
                        <EmptyDataTemplate>
                            <div class="text-center">No record found</div>
                        </EmptyDataTemplate>
                        <Columns>
                            <asp:BoundField HeaderText="ID" DataField="ID" />
                            <asp:BoundField HeaderText="Name" DataField="Name" />
                            <asp:BoundField HeaderText="Position" DataField="Position" />
                            <asp:BoundField HeaderText="Office" DataField="Office" />
                            <asp:BoundField HeaderText="Salary" DataField="Salary" />
                        </Columns>
                    </asp:GridView>-->
                </div>
            </div>
        </div>
        </div>
        
    </form>
</body>
</html>
