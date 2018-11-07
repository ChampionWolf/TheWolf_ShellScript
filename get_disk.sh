#!/bin/bash

namespace=Windows-custom-monitor
metric_names=(Disk-Free-space-c Disk-Free-space-d)

start_time=2018-10-31T23:55:00
end_time=2018-10-31T23:59:59
period=360
for metric_name in ${metric_names[@]}
do
  aws ec2 describe-instances --region cn-north-1 --query '[Reservations[*].Instances[*].[InstanceId][]]' --output text > /root/moncpu/id.txt
  /bin/cat /root/moncpu/id.txt | while read instance_id
  do
    sum=`aws cloudwatch get-metric-statistics --namespace $namespace --metric-name $metric_name --dimensions Name=InstanceId,Value=$instance_id --statistics Maximum --start-time $start_time --end-time $end_time --period $period  --output text |  awk '{print $2}'`
    if [ ! $sum ];then
      continue
    fi
    echo $instance_id: $metric_name
    echo Usage $sum | awk '{printf("%s : %.2f\n",$1,100-$2)}'
done
done
