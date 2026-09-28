@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM Tong hop moi file trong thu muc du_lieu theo thang, quy, nam -> thu muc ket_qua
python -m thongke tong-hop du_lieu --theo thang -o ket_qua\TongHop_theo_thang.xlsx
python -m thongke tong-hop du_lieu --theo quy -o ket_qua\TongHop_theo_quy.xlsx
python -m thongke tong-hop du_lieu --theo nam -o ket_qua\TongHop_theo_nam.xlsx
pause
