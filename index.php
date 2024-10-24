<?php
require 'vendor/autoload.php';

include "cfg/dbconnect.php";

$file_err = $err_msg = $succ_msg = "";
$valid_ext = array("xls", "xlsx");
$upload_dir = "uploads/";

if (isset($_POST['submit'])) {
    if ($_FILES['input_file']['name'] == "") {
        $file_err = "Please select a file";
    } else {
        // proceed for upload
        $file_name = $_FILES['input_file']['name'];
        $tmp_name = $_FILES['input_file']['tmp_name'];

        $ext = strtolower(pathinfo($file_name, PATHINFO_EXTENSION));
        if (in_array($ext, $valid_ext)) {
            // valid file
            $new_file = time() . "-" . basename($file_name);
            $conn->begin_transaction();
            try {
                move_uploaded_file($tmp_name, $upload_dir . $new_file);
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                $spreadsheet = $reader->load($upload_dir . $new_file);
                $worksheet = $spreadsheet->getActiveSheet();
                $data = $worksheet->toArray();
                unset($data[0]); // remove header
                foreach ($data as $row) {
                    $name = $row[0];
                    $email = $row[1];
                    $age = $row[2];
                    $gender = $row[3];

                    // check if customer exists
                    $sql = "select * from customers where email = ?";
                    $stmt = $conn->prepare($sql);
                    $stmt->bind_param("s", $email);
                    $stmt->execute();
                    $result = $stmt->get_result();
                    if ($result->num_rows > 0) {
                        // customer exists, update
                        $sql = "update customers set name = ?, age = ?, gender = ? where email = ?";
                        $stmt = $conn->prepare($sql);
                        $stmt->bind_param("siss", $name, $age, $gender, $email);
                        $stmt->execute();
                    } else {
                        // new customer, insert 
                        $sql = "insert into customers (name, email, age, gender) values(?,?,?,?)";
                        $stmt = $conn->prepare($sql);
                        $stmt->bind_param("ssis", $name, $email, $age, $gender);
                        $stmt->execute();
                    }
                }
                $conn->commit();
                $succ_msg = "File Uploaded successfully";
            } catch (Exception $e) {
                $conn->rollback();
                $err_msg = $e->getMessage();
            }
        } else {
            $err_msg = "Invalid File";
        }
    }
}
?>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Import Excel data into MySQL table</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="css/style.css">
</head>

<body>
    <div class="container">
        <h1>Import Excel data into MySQL table</h1>
        <?php
        if (!empty($err_msg)) { ?>
            <div class="alert alert-danger"><?= $err_msg ?></div>
        <?php }
        if (!empty($succ_msg)) { ?>
            <div class="alert alert-success"><?= $succ_msg ?></div>
        <?php }
        ?>
        <form action="" method="post" enctype="multipart/form-data">
            <div class="mb-3">
                <label for="input_file" class="form-label fw-bold">Upload file</label>
                <input
                    type="file"
                    class="form-control"
                    name="input_file"
                    id="input_file"
                    placeholder=""
                    aria-describedby="fileHelpId" />
                <div id="fileHelpId" class="form-text">Allowed File Types: xls, xlsx. Must have header line.</div>
            </div>
            <div class="text-danger"><?= $file_err ?></div>

            <button
                type="submit"
                class="btn btn-primary" name="submit">
                Submit
            </button>
        </form>
    </div>

</body>

</html>