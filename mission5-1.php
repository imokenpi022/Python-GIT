<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>入力フォームを試す</title>
</head>
<body>
    <form action="" method="post">
        <input type="text" name="name" placeholder="名前"><br>
        <input type="text" name="comment" placeholder="コメント"><br>
        <input type="submit" name="submit">
    </form>
    
    <?php




    // DB接続設定
    $dsn = 'mysql:dbname=データベース名;host=localhost';
    $user = 'ユーザ名';
    $password = 'パスワード';
    $pdo = new PDO($dsn, $user, $password, array(PDO::ATTR_ERRMODE => PDO::ERRMODE_WARNING));
 
    
    

    
      if( !empty( $_POST["name"] ) && !empty($_POST["comment"])){
          
        
    




    
    $name = $_POST["name"];
    $comment = $_POST["comment"]; 
    $date = date ( "Y/m/d H:i:s" );

    $sql = "INSERT INTO tbtest (name, comment, date) VALUES (:name, :comment, :date)";
    $stmt = $pdo->prepare($sql);
    //プレースホルダに変数を宛がう
    $stmt->bindParam(':name', $name, PDO::PARAM_STR);
    $stmt->bindParam(':comment', $comment, PDO::PARAM_STR); 
    $stmt->bindParam(':date', $date, PDO::PARAM_STR); 


    //実行
    $stmt->execute();

      }
   //$rowの添字（[ ]内）は、4-2で作成したカラムの名称に合わせる必要があります。
   $sql = 'SELECT * FROM tbtest';
   $stmt = $pdo->query($sql);
   $results = $stmt->fetchAll(); 

   //ループして、取得したデータを表示
   foreach ($results as $row){
       //$rowの中にはテーブルのカラム名が入る
       echo $row['id'].',';
       echo $row['name'].',';
       echo $row['comment'].',';
       echo $row['date'].'<br>';
   echo "<hr>";
   }
 





   
    ?>
</body>
</html>
