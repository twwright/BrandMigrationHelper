git status
git add mysite/*
rm -f .git/index.lock
git add mysite/*
git add .
git status
git commit "Update line items"
git commit -m "Add yes/no flow to series output"
git push
git status
git reset HEAD~1
git stash
git pull --ff-only
git stash apply
git status
git add .
git commit -m "Add yes/no flow to series output"
git push
pwd
ls
cd mysite
ls
cp series.py customer_series.py
ls
git add .
git commit -m "New file for sold series"
git push
exit
