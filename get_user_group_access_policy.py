import oci
import datetime
import xlwt
import re,sys

def list_domain_users(domain):
    identity_domains_client = oci.identity_domains.IdentityDomainsClient(config, domain.url)
    users = []
    next_page = None
    while True:
        response = identity_domains_client.list_users(
            attribute_sets=['all'],
            limit=100,
            page=next_page 
        )
        for user in response.data.resources:
            user_groups = []
            if user.groups is not None:
                for group in user.groups:                    
                    user_groups.append(group.display)

            user_info = {
                'domain_display_name': domain.display_name,
                'domain_url': domain.url,
                'display_name': user.display_name,
                'user_name': user.user_name,
                'user_ocid': user.ocid,
                'groups': user_groups
            }
            users.append(user_info)
        if not response.has_next_page:
            break
        next_page = response.next_page
    return users

def get_filtered_policies(policies):
    def format_group_name(statement):
        match = re.search(r'allow group (\S+)', statement)
        if match:
            group_name = match.group(1)
            if '/' in group_name:
                domain, group = group_name.split('/', 1)
                if not (domain.startswith("'") and domain.endswith("'")):
                    domain = f"'{domain}'"
                if not (group.startswith("'") and group.endswith("'")):
                    group = f"'{group}'"
                formatted_group = f"{domain}/{group}"
            else:
                if not (group_name.startswith("'") and group_name.endswith("'")):
                    formatted_group = f"'default'/'{group_name}'"
                else:
                    formatted_group = f"'default'/{group_name}"
            return statement.replace(group_name, formatted_group)
        return statement

    policy_info_list = []
    for policy in policies.data:
        filtered_statements = []
        for statement in policy.statements:
            statement_lower = statement.lower()
            if 'allow group' in statement_lower:
                filtered_statements.append(format_group_name(statement_lower))
            elif 'any-user' in statement_lower or 'any-group' in statement_lower:
                filtered_statements.append(statement_lower)

        if filtered_statements:
            policy_info_list.append({
                'policy_name': policy.name,
                'filtered_statements': filtered_statements
            })
    return policy_info_list

def get_user_policies(user_info, policies):
    domain_display_name = user_info['domain_display_name']
    groups = user_info['groups']
    
    results = []
    for policy in policies:
        policy_name = policy['policy_name']
        filtered_statements = policy['filtered_statements']
        
        for statement in filtered_statements:
            if 'any-user' in statement or 'any-group' in statement:
                results.append({
                    'group_name': 'any-user/any-group',
                    'policy_name': policy_name,
                    'statement': statement
                })
            else:
                for group in groups:
                    formatted_group_name = f"'{domain_display_name}'/'{group}'".lower()
                    if formatted_group_name in statement:
                        results.append({
                            'group_name': group,
                            'policy_name': policy_name,
                            'statement': statement
                        })
    
    return results

config = oci.config.from_file()
tenancy_id = config["tenancy"]

identity = oci.identity.IdentityClient(config)
identity_client = oci.identity.IdentityClient(config)

# Create a new Excel workbook and worksheet for user policies
wb_users_policies = xlwt.Workbook()
ws_users_policies = wb_users_policies.add_sheet("User Policies")

# Write the header row for user policies
headers_users_policies = ["Domain Display Name", "Display Name", "User Name", "User OCID", "Group", "Policy Name", "Statement"]
for col, header in enumerate(headers_users_policies):
    ws_users_policies.write(0, col, header)

now = datetime.datetime.now()
print("Start time:", now)

# Fetch all users and their policies before processing audit events
all_users = []
domains = identity_client.list_domains(tenancy_id).data
for domain in domains:
    if domain.lifecycle_state == 'ACTIVE':
        print(f"Fetching users for domain: {domain.display_name}")
        domain_users = list_domain_users(domain)
        all_users.extend(domain_users)
    else:
        print(f"Skip domain {domain.display_name}: Lifecycle - {domain.lifecycle_state}")

policies = identity_client.list_policies(tenancy_id)
policy_info_list = get_filtered_policies(policies)

#get user policies
row_num_users_policies = 1
for user in all_users:
    user_policies = get_user_policies(user, policy_info_list)
    if user_policies:
        for policy in user_policies:
            ws_users_policies.write(row_num_users_policies, 0, user['domain_display_name'])
            ws_users_policies.write(row_num_users_policies, 1, user['display_name'])
            ws_users_policies.write(row_num_users_policies, 2, user['user_name'])
            ws_users_policies.write(row_num_users_policies, 3, user['user_ocid'])
            ws_users_policies.write(row_num_users_policies, 4, policy['group_name'])
            ws_users_policies.write(row_num_users_policies, 5, policy['policy_name'])
            ws_users_policies.write(row_num_users_policies, 6, policy['statement'])
            row_num_users_policies += 1

excel_file_events = r"C:\Security\Blogs\Access_Analyzer\logs\user_policies.xls"
wb_users_policies.save(excel_file_events)

now = datetime.datetime.now()
print("End time:", now)
