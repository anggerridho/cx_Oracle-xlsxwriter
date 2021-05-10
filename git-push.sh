echo "Backup XLSX & CSV"
mv /root/Daily/*.xlsx /root/Backup_Daily/
mv /root/Daily/*.csv /root/Backup_Daily/

cd /root/Daily/
echo "# Execute: git status"
git status
echo ""
#git add RBT_Daily_Report.ipynb
echo "# Execute: git add ."
git add .
echo ""
echo "# Execute: git commit -m "RBT_Daily_Report""
git commit -m "RBT_Daily_Report"
echo ""
echo "# Execute: git status"
git status
echo ""
echo "# Execute: git push -u origin anggerridho"
git push -u origin anggerridho
echo ""

echo "Restore XLSX & CSV"
mv /root/Backup_Daily/*.xlsx /root/Daily/
mv /root/Backup_Daily/*.csv /root/Daily/

# Note if error
# https://stackoverflow.com/questions/24114676/git-error-failed-to-push-some-refs-to-remote
#git pull --rebase
