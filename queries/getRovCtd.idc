Datasource: expd_fog
SQLStatement: select * 
+from RovCtd
+ where usec > '%start%' and usec < '%end%'
Template: getRovCtd.htx
Username: expddba
Password: password
