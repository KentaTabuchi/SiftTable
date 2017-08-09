cd bin
move staff_info.accdb staff_info.dat
cd ../
cscript vbac.wsf decombine
cd bin
move staff_info.dat staff_info.accdb
pause