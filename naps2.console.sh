#!/bin/bash
echo
echo -n "Fake NAPS2 console.  Args: "
echo $1 $2
sleep 10
touch $2
echo -n "Created 0-byte pdf: "
echo $2

