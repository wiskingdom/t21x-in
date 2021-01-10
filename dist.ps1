Copy-Item -Path .\for-dist\manifest.xml -Destination .\t21x-in\ -Force
npm run build
firebase deploy