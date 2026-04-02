#!/bin/bash

cp ~/Documents/Codex_File/Blind\ Test/index.html ~/Documents/GitHub/EOF-Script-blind-test-dashboard/index.html

cp ~/Documents/Codex_File/Blind\ Test/blind-test-dashboard-data-0401.js ~/Documents/GitHub/EOF-Script-blind-test-dashboard/blind-test-dashboard-data-0401.js

cp ~/Documents/Codex_File/Blind\ Test/recomputed_rows_raw.json ~/Documents/GitHub/EOF-Script-blind-test-dashboard/recomputed_rows_raw.json

cp ~/Documents/Codex_File/Blind\ Test/workbook_rows_0401.json ~/Documents/GitHub/EOF-Script-blind-test-dashboard/workbook_rows_0401.json

cd ~/Documents/GitHub/EOF-Script-blind-test-dashboard
./auto_push.sh
