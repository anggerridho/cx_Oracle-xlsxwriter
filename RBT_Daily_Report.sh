#!/bin/bash
export ORACLE_HOME=/usr/lib/oracle/12.2/client64
export LD_LIBRARY_PATH=$ORACLE_HOME/lib:$LD_LIBRARY_PATH
>/root/LOG/RBT_Daily_Report.log
/root/anaconda3/bin/python -u /root/Daily/RBT_Daily_Report.py | tee -a /root/LOG/RBT_Daily_Report.log
/root/trxdaily.sh > /root/LOG/trxdaily.log
