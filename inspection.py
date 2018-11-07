import sys
import sqlite3
import xlsxwriter

from boto3 import Session


access_key = "3123safsfdsfsdf"
secret_key = "fsadfasdf2314jk"

session = Session(access_key, secret_key)


'''
pip install xlsxwriter
'''

# '''创建一个数据库，文件名'''
# conn = sqlite3.connect('aws.db')
# print("Opened database successfully")


# sys.setdefaultencoding("utf-8")
# path = os.path.dirname(os.path.abspath(__file__))

# 建立文件
workbook = xlsxwriter.Workbook("巡检表.xlsx")
# 可以制定表的名字
# worksheet = workbook.add_worksheet('text')

worksheet = workbook.add_worksheet('平台资源统计')
worksheet2 = workbook.add_worksheet('账单详情')
worksheet3 = workbook.add_worksheet('实例资源详情')
worksheet4 = workbook.add_worksheet('端口安全')
worksheet5 = workbook.add_worksheet('异常&建议')

# 设置列宽
worksheet.set_column('A:A', 5.83)
worksheet.set_column('B:B', 16.17)
worksheet.set_column('C:C', 21.67)
worksheet.set_column('D:D', 12.17)
worksheet.set_column('E:E', 51.17)

# 设置祖体
bold = workbook.add_format({'bold': True})
# 定义数字格式
# money = workbook.add_format({'num_format':'$#,##0'})

# border：边框，align:对齐方式，bg_color：背景颜色，font_size：字体大小，bold：字体加粗，font_color：字体颜色，font_name：字体类型
top = workbook.add_format({
    'border': 1,
    'align': 'center',
    'bg_color': '#696969',
    'font_name': 'Arial',
    'font_size': 14,
    'font_color': 'white',
    # 'bold': True
})
# 
security = workbook.add_format({
    'border': 1,
    'align': 'center',
    'bg_color': '#DCDCDC',
    'font_name': 'Arial',
    'font_size': 14,
    # 'font_color': 'white',
    'bold': True
})
# 合并参数
merge_format = workbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',  # 垂直居中
    # 'bg_color': 'DimGray',
    'font_name': 'Arial',
    'font_size': 14,
    # 'font_color': 'white',
    # 'bold': True
})
merge_format_security = workbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',  # 垂直居中
    'bg_color': '#696969',
    'font_name': 'Arial',
    'font_size': 14,
    'font_color': 'white',
    # 'bold': True
})
# 内容
content = workbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',  # 垂直居中
    'font_name': 'Arial',
    'font_size': 14,
})


# 端口安全
security_title = [u'安全组ID', u'作用域(实例)', u'协议', u'开放端口', u'允许网段', u'合规性']
# 平台资源统计
resource_title = [u'NO.', u'Type', u'Name', u'Total', u'Description']

# 写入带粗体的数据
worksheet.write_row('A1', resource_title, top)

worksheet4.merge_range('B1:G1', '当前安全组信息整理', merge_format_security)
# worksheet4.write('B2', '安全组ID.', security)

worksheet4.write_row('B2', security_title, security)

# 合并
worksheet.merge_range('B2:B6', 'VPC', merge_format)
worksheet.merge_range('B7:B13', 'EC2', merge_format)
worksheet.merge_range('B14:B16', 'IAM', merge_format)
worksheet.merge_range('B17:B19', 'CloudWatch', merge_format)

# 0，0行，列
worksheet.write(1, 2, 'VPC', content)     # write_string()
worksheet.write(2, 2, 'Subents', content)     # write_string()
worksheet.write(3, 2, 'Routers', content)        # write_number()
worksheet.write(4, 2, 'IGW', content)     # write_number()
worksheet.write(5, 2, 'Security Groups', content)  # write_formula()
worksheet.write(6, 2, 'Instances', content)        # write_blank()
# TheWolf Add
worksheet.write(7, 2, 'Reserved Instance', content)       # write_blank()

worksheet.write(8, 2, 'AMI', content)       # write_blank()
worksheet.write(9, 2, 'Volumes Number', content)       # write_blank()
worksheet.write(10, 2, 'Volumes Size', content)       # write_blank()
worksheet.write(11, 2, 'EIP', content)       # write_blank()
worksheet.write(12, 2, 'Key pairs', content)       # write_blank()
worksheet.write(13, 2, 'Groups', content)       # write_blank()
worksheet.write(14, 2, 'Users', content)       # write_blank()
worksheet.write(15, 2, 'Roles', content)       # write_blank()
worksheet.write(16, 2, 'Default Metrics', content)       # write_blank()
worksheet.write(17, 2, 'Customize Metrics', content)       # write_blank()
worksheet.write(18, 2, 'Alarms', content)       # write_blank()


# 第三列

worksheet.write(16, 3, '', content)       # write_blank()


# 第四列
worksheet.write(1, 4, 'virtual private cloud', content)       # write_blank()
worksheet.write(2, 4, 'Subents', content)       # write_blank()
worksheet.write(3, 4, 'Routers Table', content)       # write_blank()
worksheet.write(4, 4, 'Internet Gateway', content)       # write_blank()
worksheet.write(5, 4, 'Working As Virtual Network Firewall', content)       # write_blank()
worksheet.write(6, 4, 'Elastic Compute Cloud', content)       # write_blank()
# TheWolf Add
worksheet.write(7, 4, 'EC2 Reserved Instance', content)       # write_blank()

worksheet.write(8, 4, 'Amazon System Image', content)       # write_blank()
worksheet.write(9, 4, 'Elastic Block Store', content)       # write_blank()
worksheet.write(10, 4, 'Total Size Of EBS (GB)', content)       # write_blank()
worksheet.write(11, 4, 'Elastic IP address', content)       # write_blank()
worksheet.write(12, 4, 'Public And Private Keys ', content)       # write_blank()
worksheet.write(13, 4, 'Collections of IAM users', content)       # write_blank()
worksheet.write(14, 4, 'IAM identities', content)       # write_blank()
worksheet.write(15, 4, 'Identities with permission policies', content)       # write_blank()
worksheet.write(16, 4, 'Amazon CloudWatch Metrics For All', content)       # write_blank()
worksheet.write(17, 4, 'Amazon CloudWatch Metrics For Customer', content)       # write_blank()
worksheet.write(18, 4, 'Amazon CloudWatch Items For All', content)       # write_blank()

def get_target_value(key, dic, tmp_list):
    """
    :param key: 目标key值
    :param dic: JSON数据
    :param tmp_list: 用于存储获取的数据
    :return: list
    """
    if not isinstance(dic, dict) or not isinstance(tmp_list, list):  # 对传入数据进行格式校验
        return 'argv[1] not an dict or argv[-1] not an list '

    if key in dic.keys():
        tmp_list.append(dic[key])  # 传入数据存在则存入tmp_list
    else:
        for value in dic.values():  # 传入数据不符合则对其value值进行遍历
            if isinstance(value, dict):
                get_target_value(key, value, tmp_list)  # 传入数据的value值是字典，则直接调用自身
            elif isinstance(value, (list, tuple)):
                _get_value(key, value, tmp_list)  # 传入数据的value值是列表或者元组，则调用_get_value
    return tmp_list


def _get_value(key, val, tmp_list):
    for val_ in val:
        if isinstance(val_, dict):
            get_target_value(key, val_, tmp_list)  # 传入数据的value值是字典，则调用get_target_value
        elif isinstance(val_, (list, tuple)):
            _get_value(key, val_, tmp_list)   # 传入数据的value值是列表或者元组，则调用自身


def account():
    client = session.client('sts')
    account_id = client.get_caller_identity()
    return account_id['Account']


def ec2():
    client = session.client('ec2')
    response = client.describe_vpcs(
    )
    vpc_count = 0
    for vpc in response['Vpcs']:
        if not vpc['VpcId']:
            vpc_count = 0
        else:
            vpc_count = vpc_count + 1

    # VPC数量
    worksheet.write(1, 3, vpc_count, content)  # write_blank()
    # Subents
    describe_subnets = client.describe_subnets(
        # Filters=[
        #     {
        #         'Name': 'string',
        #         'Values': [
        #             'string',
        #         ]
        #     },
        # ],
    )
    subnet_count = 0
    for Subnet in describe_subnets['Subnets']:
        if not Subnet['SubnetId']:
            subnet_count = 0
        else:
            subnet_count = subnet_count + 1

    # 子网
    worksheet.write(2, 3, subnet_count, content)  # write_blank()

    # 路由表
    describe_route_tables = client.describe_route_tables(
        # Filters=[
        #     {
        #         'Name': 'string',
        #         'Values': [
        #             'string',
        #         ]
        #     },
        # ],
    )
    route_count = 0
    for route in describe_route_tables['RouteTables']:
        if not route['RouteTableId']:
            route_count = 0
        else:
            route_count = route_count + 1
    # 路由表
    worksheet.write(3, 3, route_count, content)  # write_blank()
    # igw
    describe_internet_gateways = client.describe_internet_gateways(
        # Filters=[
        #     {
        #         'Name': 'string',
        #         'Values': [
        #             'string',
        #         ]
        #     },
        # ],
    )
    internet_gateways_count = 0
    for internet_gateways in describe_internet_gateways['InternetGateways']:
        if not internet_gateways['InternetGatewayId']:
            internet_gateways_count = 0
        else:
            internet_gateways_count = internet_gateways_count + 1
    #         igw
    worksheet.write(4, 3, internet_gateways_count, content)  # write_blank()
    # 安全组
    describe_security_groups = client.describe_security_groups()
    security_groups_count = 0
    for security_groups in describe_security_groups['SecurityGroups']:
        if not security_groups['GroupId']:
            security_groups_count = 0
        else:
            security_groups_count = security_groups_count + 1
    worksheet.write(5, 3, security_groups_count, content)  # write_blank()
    # instances
    describe_instances = client.describe_instances()
    instances_count = 0
    for instances in describe_instances['Reservations']:
        for InstanceId in instances['Instances']:

            if not InstanceId['InstanceId']:
                instances_count = 0
            else:
                instances_count = instances_count + 1
    worksheet.write(6, 3, instances_count, content)  # write_blank()
    # Reserved instance
    filters = [
        {
            'Name': 'state',
            'Values':
                ['active',
            ]
        }
    ]

    reserved_total = []
    response = client.describe_reserved_instances(Filters=filters)

    for reserved_instance in response['ReservedInstances']:
        reserved_total.append(reserved_instance['InstanceCount'])
    reserved_instance_count = sum(reserved_total)
    worksheet.write(7, 3, reserved_instance_count, content)  # write_blank()

    # AMI 默认会列出所有公共镜像
    owners = account()
    describe_images = client.describe_images(
        Owners=[
            owners,
        ],
    )
    ami_count = 0
    for images in describe_images['Images']:
        if not images['ImageId']:
            ami_count = 0
        else:
            ami_count = ami_count + 1
    worksheet.write(8, 3, ami_count, content)  # write_blank()
    # Volumes Number
    describe_volumes = client.describe_volumes(
        # Filters=[
        #     {
        #         'Name': 'string',
        #         'Values': [
        #             'string',
        #         ]
        #     },
        # ],
        # VolumeIds=[
        #     'string',
        # ],
    )
    volumes_count = 0
    volumes_size_count = 0
    for volumes in describe_volumes['Volumes']:
        if not volumes['VolumeId']:
            volumes_count = 0
        else:
            volumes_count = volumes_count + 1
        if not volumes['Size']:
            volumes_size_count = 0
        else:
            volumes_size_count = volumes_size_count + volumes['Size']
    # 卷数量
    worksheet.write(9, 3, volumes_count, content)  # write_blank()
    # 卷大小相加
    worksheet.write(10, 3, volumes_size_count, content)  # write_blank()
    # EIP
    describe_addresses = client.describe_addresses()
    addresses_count = 0
    for addresses in describe_addresses['Addresses']:
        if not addresses['AllocationId']:
            addresses_count = 0
        else:
            addresses_count = addresses_count + 1
    worksheet.write(11, 3, addresses_count, content)  # write_blank()
    # Describes one or more of your key pairs.
    describe_key_pairs = client.describe_key_pairs()
    key_pairs = 0
    for describe_key in describe_key_pairs['KeyPairs']:
        if not describe_key['KeyFingerprint']:
            key_pairs = 0
        else:
            key_pairs = key_pairs + 1
    worksheet.write(12, 3, key_pairs, content)  # write_blank()
    # security
    owners = owners.split(" ")
    describe_security_groups = client.describe_security_groups()

    groups = []
    ip_permissions = []
    ip = []
    # worksheet4.merge_range(5, 1, 2, 1, '', merge_format)
    for security_groups in describe_security_groups['SecurityGroups']:
        # 获取GroupId,存入列表
        groups.append(security_groups['GroupId'])
    worksheet4.write_column('B3', groups, content)

def iam():
    client = session.client('iam')
    list_groups = client.list_groups(
    )
    grops_count = 0
    for grop in list_groups['Groups']:
        if not grop['Arn']:
            grops_count = 0
        else:
            grops_count = grops_count + 1
    worksheet.write(13, 3, grops_count, content)  # grops
    # iam user
    list_users = client.list_users(
    )
    users_count = 0
    for users in list_users['Users']:
        if not users['UserId']:
            users_count = 0
        else:
            users_count = users_count + 1
    worksheet.write(14, 3, users_count, content)  # write_blank()
    # iam roles
    list_roles = client.list_roles()
    roles_count = 0
    for roles in list_roles['Roles']:
        if not roles['RoleId']:
            roles_count = 0
        else:
            roles_count = roles_count + 1
    worksheet.write(15, 3, roles_count, content)  # write_blank()


def cloudwatch():
    client = session.client('cloudwatch')
    cloudwatch = session.resource('cloudwatch')

    default_metric = 0
    aws_metric = 0
    metric_iterator = cloudwatch.metrics.all()
    for metric in metric_iterator:
        metric = str(metric)
        if 'AWS/' in metric:
            default_metric += 1
        else:
            aws_metric += 1

    worksheet.write(16, 3, default_metric, content)  # write_blank()
    worksheet.write(17, 3, aws_metric, content)  # write_blank()


    # alarms
    describe_alarms = client.describe_alarms()
    alarms_count = 0
    for alarms in describe_alarms['MetricAlarms']:
        if not alarms['AlarmName']:
            alarms_count = 0
        else:
            alarms_count = alarms_count + 1
    worksheet.write(18, 3, alarms_count, content)  # write_blank()

if __name__ == '__main__':
    ec2()
    iam()
    cloudwatch()
    workbook.close()
