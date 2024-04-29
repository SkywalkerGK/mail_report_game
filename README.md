### Crontab Mail Report Game ###
###  start csv auto report game bybass  ###

00 8 1 * * sh /home/mapuser/report_game_bybass/auto_report_game_bybass.sh >> /home/mapuser/report_game_bybass/log/auto_log_report_game_bybass.log;
59 23 31 12 * sh /home/mapuser/report_game_bybass/clear_tmp_report_game_bybass.sh;

###  end csv auto report game bybass  ###
