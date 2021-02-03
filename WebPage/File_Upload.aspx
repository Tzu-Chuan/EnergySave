<%@ Page Language="C#" AutoEventWireup="true" CodeFile="File_Upload.aspx.cs" Inherits="WebPage_File_Upload" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="../App_Themes/css/OchiLayout.css" rel="stylesheet" type="text/css" /><!-- ochsion layout base -->
    <link href="../App_Themes/css/OchiColor.css" rel="stylesheet" type="text/css" /><!-- ochsion layout color -->
    <link href="../App_Themes/css/jquery-ui.css" rel="stylesheet" />
    <script src="../js/jquery-1.11.2.min.js"></script>
    <script src="../js/jquery-ui-1.10.2.custom.min.js"></script>
    <script src="../js/downfile.js"></script>
    <script src="../js_plupload/plupload.full.min.js"></script>
    <title></title>
    <script type="text/javascript">
        $(document).ready(function () {
            try {
                //function CreateGuid() {
                //    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
                //        var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
                //        return v.toString(16);
                //    });
                //}
                var uploader = new plupload.Uploader({
                    runtimes: 'html5,flash,silverlight,html4',
                    browse_button: 'pickfiles', // you can pass in id...
                    container: document.getElementById('uploadblock'), // ... or DOM Element itself
                    url: '../Handler/uploadFile.ashx',
                    flash_swf_url: '../js_plupload/Moxie.swf',
                    silverlight_xap_url: '../js_plupload/Moxie.xap',
                    multipart_params: {
                        pGuid: $.getParamValue('v'),
                        type: $.getParamValue('tp')
                    },
                    filters: {
                        // Maximum file size
                        max_file_size: '2000mb',
                        // Specify what files to browse for
                        mime_types: [
                            { title: "all files", extensions: "*" }
                        ]
                    },
                    init: {
                        Init: function (up, params) {
                            $('#warningStr').html("<div>目前瀏覽器優先使用: " + params.runtime + "</div>");
                        },
                        FilesAdded: function (up, files) {
                            $("#recStr").html("");
                            if (up.total.size >= 2147483648) {
                                up.removeFile(files[0]);
                                alert("所有檔案相加其大小不可超過2GB");
                                return;
                            }

                            //判斷副檔名
                            var extenError = false;
                            $.each(up.files, function (i, file) {
                                var filename = file.name.toLowerCase();
                                var extary = filename.split('.');
								var exten = "." + extary[extary.length - 1];
                                if (exten != ".doc" && exten != ".docx" && exten != ".xls" && exten != ".xlsx" && exten != ".ppt" && exten != ".pptx" && exten != ".pdf") {
                                    alert("檔案格式限制：.doc、.docx、.xls、.xlsx、.ppt、.pptx、.pdf");
                                    extenError = true;
                                    $.each(files, function (i, file) {
                                        up.removeFile(file);
                                    });
                                    return false;
                                }
                            });
                            if (extenError == true)
                                return false;


                            var deleteHandle = function (uploaderObject, fileObject) {
                                return function (event) {
                                    event.preventDefault();
                                    uploaderObject.removeFile(fileObject);
                                    $(this).closest("li#" + fileObject.id).remove();
                                };
                            };

                            var fileListTable = '<div name="ListTitle" style="border-bottom:solid 1px #dddddd;">檔案列表</div>';
                            if ($("#plupload_content").find("div[name='ListTitle']").length == 0) {
                                $("#plupload_content").prepend(fileListTable);
                            }

                            var Flist = '<ul style="list-style-type: none;">';
                            $.each(files, function (i, file) {
                                Flist += '<li id="' + file.id + '">';
                                Flist += '<!--<div class="plupload_fileSizeList" ><span>' + plupload.formatSize(file.size) + '</span></div > -->';
                                Flist += '<div class="fileNameList" style="padding-bottom:5px; border-bottom:solid 1px #dddddd;"><a href="javascript:void(0);"><img src="../App_Themes/images/icon-delete-new.png" border="0" id="deleteFile' + files[i].id + '" style="margin-top: 5px;" /></a>';
                                Flist += '<span style="margin-left:10px;">' + file.name + ' (' + plupload.formatSize(file.size) + ')</span>&nbsp;<div name="bar" style="width:100px; height:10px; display: inline-block;"></div>&nbsp;<span name="filepercent" style="padding-left:5px;"></span></li>';
                            });
                            Flist += '</ul>';
                            $('#filelist').append(Flist);

                            for (var i in files) {
                                $('#deleteFile' + files[i].id).click(deleteHandle(up, files[i]));
                                //調整 li 換行後的間距
                                $("li#" + files[i].id + "").attr("style", "margin-bottom:3px");
                            }
                        },
                        //BeforeUpload: function (up, files) {
                        //    $.extend(up.settings.multipart_params, { myfileid: $('#' + files.id + 'myfileid').val() });
                        //},
                        FileUploaded: function (up, files) {
                         
                        },
                        UploadComplete: function (up, files) {
                            if ((uploader.total.uploaded) == uploader.files.length) {
                                $("#recStr").html("檔案上傳成功");
                                if ($.getParamValue('tp') == "06") {
                                    parent.upFile_feedback($.getParamValue('v'));
                                }
                                else {
                                    parent.upFile_feedback($.getParamValue('tp'));
                                }
                                $("#loadsp").hide();
                                $("#btnsp").show();
                                //parent.$.fancybox.close();
                            }
                        },
                        UploadProgress: function (up, file) {
                            $("#" + file.id + "").find("span[name='filepercent']").html('<span>' + file.percent + '%</span>');
                            $("#" + file.id + "").find("div[name='bar']").progressbar({
                                value: file.percent
                            });
                        },
                        Error: function (up, err) {
                            if (err.code == "-600") {
                                alert("請選擇低於2GB的檔案上傳");
                            }
                            else {
                                $('#warningStr').append("<div>Error: " + err.code + ", Message: " + err.message + (err.file ? ", File: " + err.file.name : "") + "</div>");
                                var text = err.response;
                                var startIndex = text.indexOf("<title>");
                                var endIndex = text.indexOf("</title>");
                                var innerText = text.slice(startIndex + 7, endIndex);
                                $("#ExceptionStr").html("[Exception]:" + innerText);
                            }
                            up.refresh(); // Reposition Flash/Silverlight
                            $("#loadsp").hide();
                            $("#btnsp").show();
                        }
                    }
                });

            }
            catch (err) {
                alert(err);
                $("#loadsp").hide();
                $("#btnsp").show();
            }

            uploader.init();

            $(document).on("click", "#btnCancel", function () {
                if (confirm("確定離開?") == true) {
                    parent.$.fancybox.close();
                }
            });

            $(document).on("click", "#btnAdd", function () {
                if ($('#filelist').html().trim().length == 0) {
                    alert("請選擇檔案")
                    return false;
                }
                else {
                    $("#loadsp").show();
                    $("#btnsp").hide();
                    uploader.start();
                }
            });
        });//js end
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div class="stripeMe" id="uploadblock">
            <table width="98%" border="0" cellspacing="0" cellpadding="0">
                <tr><th colspan="2">檔案上傳</th></tr>
                <tr id="attachment">
                    <td>
                        <div id="divif">
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td>
                                        <span style="color:Red;">
                                            格式限制：.doc、.docx、.xls、.xlsx、.ppt、.pptx、.pdf<br />
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <div id="container">
                                            <input type="button" id="pickfiles" value="選擇檔案" class="genbtn" />
                                        </div>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <div id="plupload_content">
                                            <div id="filelist" style="margin-left:-20px;"></div>
                                            <div id="recStr" style="color:red;"></div>
                                        </div>
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </td>
                </tr>
                <tr>
                <td colspan="2" align="right">
                    <span id="loadsp" style="display:none;"><img alt="Loading..." src="../App_Themes/images/loading.gif" width="30" style="position:absolute; right:110px;" />檔案上傳中...</span>
                    <span id="btnsp">
                        <input id="btnAdd" type="button" value="新增" class="genbtn" />
                        <input id="btnCancel" type="button" value="取消" class="genbtn" />
                    </span>
                </td>
                </tr>
            </table>
        </div>
        <div class="warningStr" style="margin-top:5px;">您的瀏覽器不支援Flash,Silverlight以及HTML5</div>
        <div id="ExceptionStr" style="margin-top:5px; color:red;"></div>
    </div>
    </form>
</body>
</html>
