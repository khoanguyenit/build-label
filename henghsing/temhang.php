<?php session_start();?>
<?php if(isset($_POST['name'])) {$_SESSION['name'] = $_POST['name'];}?>
<?php 
	if(isset($_SESSION['name']) && $_SESSION['name'] != 'henghsing'){
		echo "<script>alert('Dang nhap bat hop phap');</script>";
		echo "<script>location.href='http://www.google.com.vn'</script>";
	}
?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="style.css" rel="stylesheet" type="text/css" />
    <title>Upload File Excel</title>
</head>
<style type="text/css">
    .clearfix {
        display: block;
        content: "";
        clear: both;
    }
    #header{
        margin-top:50px;
        width:100%;
        height:70px;
        background:#265CFF;
        color:#FFF;
        text-align:center;
        font-size:22pt;
        padding-top:30px;
    }
    #hethong{
        width:500px;
        margin:auto;
        margin-top:200px;
    }
    input[type="submit"]{
        background:#265CFF;
        padding:10px 20px;
        border:none;
        color:#FFF;
        -webkit-border-radius: 5px;
        -moz-border-radius: 5px;
        border-radius: 5px;
    }
    #footer{
        margin-top:200px;
        text-align:center;
        color:#265CFF;
        border-bottom:1px solid #265CFF;
        border-top:1px solid #265CFF;
        padding-top:20px;
        padding-bottom:20px;
    }
    #thongbao{
        display:none;
        color:#265CFF;
    }
</style>
<script>
    function check(){
        document.getElementById('thongbao').style.display = 'block';
    }
</script>
<body>
    <div id="header">
        Phầm Mềm Xuất Tem Hàng Heng Hsing Ver 2.0
    </div>
    <div id="hethong">
    <form enctype="multipart/form-data" action="xuly.php" method="post" onsubmit="check()">
        <input type="hidden" name="MAX_FILE_SIZE" value="2000000" />
        <table id="table00" width="500px">
            <tr>
                <td>
                    <strong>File Excel</strong>:</td>
                <td>
                    <input type="file" name="file" required />
                </td>
                <td>
                    <input type="submit" value="Xử Lý" />
                </td>
            </tr>
	     <tr>
                <td>
                    <strong>Số Lượng / Thùng</strong>:</td>
                <td>
                    <input type="number" name="number" required />
                </td>
                <td></td>
            </tr>

            <tr style="height:30px;"></tr>
        </table>
    </form>
    <div id="thongbao">Xin vui lòng chờ hệ thống xử lý...</div>
    </div>
    <div id="footer">
        Copyright 2015 &copy; <a href="mailto:khoanguyen.it@hotmail.com">Khoa Nguyễn</a>. All Right Reserved. 
    </div>
</body>

</html>
