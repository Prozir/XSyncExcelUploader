<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UploadExcel.aspx.cs" Inherits="XSyncExcelUploader.UploadExcel" UICulture="en" Culture="en-US" %>

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

        function ShowProgress() {
            var divDetails = document.getElementById('progressdiv');
            divDetails.style.display = "block";
        }
    </script>
    <script type="text/javascript">       
        function HideProgress() {
            var divDetails = document.getElementById('progressdiv');
            divDetails.style.display = "none";
        }
    </script>

    <style type="text/css">
        .loadingmodal {
            position: fixed;
            top: 0;
            left: 0;
            background-color: black;
            z-index: 99;
            opacity: 0.8;
            filter: alpha(opacity=80);
            -moz-opacity: 0.8;
            min-height: 100%;
            width: 100%;
        }

        .loadingdiv {
            font-family: Arial;
            font-size: 10pt;
            border: 1px solid red;
            width: 30%;
            height: 100px;
            position: fixed;
            background-color: White;
            margin-left: 10%;
            z-index: 999;
        }

        .removemodalclass {
            display: none;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div id="mastercontainer" style="background: url(https://thalappakatti.com/wp-content/themes/thalappakatti/images/banner2.png); height: 100vh;">
            <div class="container py-3" style="padding-top: 4% !important">
                <div style="margin-left: 39%">
                    <img src="https://thalappakatti.com/wp-content/uploads/2018/01/thalappakatti-logo-anim.gif" rel="logo" alt="Dindigul Thalappakatti Restaurant"></div>
                <br />
                <h2 style="color: #ffbd23" class="text-center text-uppercase">Thalappakatti Excel Uploader</h2>
                <div class="card">
                    <div style="background-color: #690514 !important" class="card-header bg-primary text-uppercase text-white">
                        <h5 style="color: #ffbd23">Select Your Excel File</h5>
                    </div>
                    <div class="card-body">
                        <button style="margin-bottom: 10px; background-color: #690514 !important; color: #ffbd23" type="button" class="btn btn-primary" data-toggle="modal" data-target="#myModal">
                            <i class="fa fa-plus-circle"></i>Add Excel
                        </button>
                        <br />
                        <asp:Label ID="lblMessage" runat="server"></asp:Label>


                        <div class="modal fade" id="myModal">
                            <div class="modal-dialog">
                                <div class="modal-content" style="width: 90%">
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
                                                        </div>
                                                        <label id="filename"></label>
                                                        <div class="input-group-append">
                                                            <asp:Button ID="btnUpload" runat="server" BackColor="#690514" ForeColor="#ffbd23" CssClass="btn btn-outline-primary" Text="Upload" OnClick="btnUpload_Click" />
                                                        </div>
                                                        <br />
                                                        <asp:Label ID="progresslbl" Visible="false" runat="server" Text="ins.."></asp:Label>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>

                                    <div class="modal-footer">
                                        <button type="button" style="background-color: #690514 !important; color: #ffbd23 !important" class="btn btn-danger" data-dismiss="modal">Close</button>
                                    </div>
                                </div>
                            </div>
                        </div>

                    </div>
                </div>
            </div>
        </div>

    </form>
</body>
</html>
