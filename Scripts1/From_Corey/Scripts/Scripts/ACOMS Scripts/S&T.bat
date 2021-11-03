@ECHO OFF
@TITLE Map S & T Drives

net use S: \\xlwu-fs-001\ang$\Shared /PERSISTENT:YES
echo S Drive has successfully been mapped

net use T: \\xlwu-fs-002\Tyndall$ /PERSISTENT:YES
echo T Drive has successfully been mapped

