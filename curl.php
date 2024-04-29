<?php

$FILEPATH = '/home/mapuser/report_game_bybass/csv/report_game'.date("Ymd").'.xls';
#echo $DATE; #report_game20230823.csv

file_put_contents($FILEPATH, fopen("http://10.11.11.206/cacti_rep/export_excel_game.php", 'r'));
