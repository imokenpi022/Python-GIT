<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>入力フォームを試す</title>
</head>
<body>
    <form action="" method="post">
        <input type="text" name="str1" placeholder="名前"><br>
        <input type="text" name="str2" placeholder="コメント"><br>
        <input type="submit" name="submit">
    </form>
    
    <?php
    
      if( !empty( $_POST["str1"] ) && !empty($_POST["str2"])){
          
          $filename = 'mission3-3.csv';
          $count = count( file( $filename ) );
          $count = $count+1;
          $str1 = $_POST["str1"];
          $str2 = $_POST["str2"];
          $date = date ( "Y/m/d H:i:s" );

     // 元となるデータを配列で用意（1件）
     $user = [
        'id' => $count,
        'name' => "$str1",
        'comment' => "$str2",
        'date' => "$date"
    ];

     // CSVファイル名
    $filename = 'mission3-3.csv';
    //ファイルを'a'モードで開く（追記）
    $fp = fopen( $filename, 'a');
    
    //csvフォーマットで書き出しする
    fputcsv($fp, $user);
    
        //ファイルを閉じる
    fclose($fp);
    
      // CSVファイル名    
    $filename = "mission3-3.csv" ;
    // member.csvを読み込む
    $fp = fopen( $filename , "r");

//ループ前：テーブルタグを作成し、テーブルヘッダーで見出しを作る
    echo '<table border="1">
          <tr>
          <th>ID</th>
          <th>Name</th>
          <th>Comment</th>
          <th>Date</th>
          </tr>'; 

    //ループ：while文でCSVファイルのデータを1つずつ繰り返し読み込む
    while($data = fgetcsv($fp)){
        // テーブルセルに配列の値を格納
        echo '<tr>';
        echo '<td>'.$data[0].'</td>';
        echo '<td>'.$data[1].'</td>';
        echo '<td>'.$data[2].'</td>';
        echo '<td>'.$data[3].'</td>';
        echo '</tr>';
    } 

    //ループ後：テーブルの閉じタグを付ける
    echo '</table>';
    
        //ファイルを閉じる
    fclose($fp);
}
  

   
    ?>
</body>
</html>