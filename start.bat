cd myapp
cmd /C "mvn clean package"
cd ..
java -jar myapp\target\myapp-1.0-SNAPSHOT.jar
pause