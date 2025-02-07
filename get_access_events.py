import datetime
import oci
import xlwt
import sys

def get_subscription_regions(identity, tenancy_id):
    list_of_regions = []
    list_regions_response = identity.list_region_subscriptions(tenancy_id)
    for r in list_regions_response.data:
        list_of_regions.append(r.region_name)
    return list_of_regions

def get_compartments(identity, tenancy_id):
    list_compartments_response = oci.pagination.list_call_get_all_results(
        identity.list_compartments,
        compartment_id=tenancy_id).data

    compartment_ocids = [c.id for c in filter(lambda c: c.lifecycle_state == 'ACTIVE', list_compartments_response)]
    return compartment_ocids

def get_audit_events(audit, compartment_ocids, start_time, end_time):
    list_of_audit_events = []
    for c in compartment_ocids:
        list_events_response=list_audit_events(audit,c, start_time, end_time)

        list_of_audit_events.extend(list_events_response)
    return list_of_audit_events

def list_audit_events(audit,compartment_id, start_time, end_time):
    events = []
    next_page = None
    while True:
        response = audit.list_events(
            compartment_id=compartment_id,
            start_time=start_time,
            end_time=end_time,
            page=next_page
        )
        events.extend(response.data)
        if not response.has_next_page:
            break
        next_page = response.next_page
    return events

def format_audit_event(region, event, user_groups):
    event_data = event.data
    return {
        "Region": region,
        "Compartment ID": event_data.compartment_id,
        "Compartment Name": event_data.compartment_name,
        "Event Name": event_data.event_name,
        "Auth Type": event_data.identity.auth_type if event_data.identity else None,
        "Principal ID": event_data.identity.principal_id if event_data.identity else None,
        "Principal Name": event_data.identity.principal_name if event_data.identity else None,
        "Event Type": event.event_type,
        "User Groups": str(user_groups['groups'])
    }

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
            user_groups=[]
            if user.groups is not None:
                for group in user.groups:                    
                    user_groups.append(group.display)

            user_info = {
                'domain_display_name': domain.display_name,
                'domain_url': domain.url,
                'display_name': user.display_name,
                'user_name': user.user_name,
                'user_ocid':user.ocid,
                'groups': user_groups
            }
            users.append(user_info)
        if not response.has_next_page:
            break
        next_page = response.next_page
    return users

def get_user_by_ocid(users_list, target_ocid):
    for user in users_list:
        if user['user_ocid'] == target_ocid:
            return user
    return user


config = oci.config.from_file()
tenancy_id = config["tenancy"]

identity = oci.identity.IdentityClient(config)
identity_client = oci.identity.IdentityClient(config)
end_time = datetime.datetime.utcnow()
start_time = end_time + datetime.timedelta(days=-1)
regions = get_subscription_regions(identity, tenancy_id)
compartments = get_compartments(identity, tenancy_id)
audit = oci.audit.AuditClient(config)

# Create a new Excel workbook and worksheet
wb = xlwt.Workbook()
ws = wb.add_sheet("Audit Events")

# Write the header row
headers = ["Region", "Compartment ID", "Compartment Name", "Event Name", "Auth Type", "Principal ID", "Principal Name", "Event Type", "User Groups"]
for col, header in enumerate(headers):
    ws.write(0, col, header)

# Set to track unique events
unique_events = set()

# Write each event directly to the Excel file
now = datetime.datetime.now()
print("Start time:", now)




#since there is not way to find Identity Domain info from user OCID, unless query all ID's and filter..we are going to get the user list and do compare
#section 1 - Get list of all users and their associated group membership info
all_users = []
domains = identity_client.list_domains(tenancy_id).data
for domain in domains:
    try:
        print(f"Fetching users for domain: {domain.display_name}")
        domain_users = list_domain_users(domain)
        all_users.extend(domain_users)
    except Exception as e:
        print(f"Failed to fetch users for domain {domain.display_name}: {e}")

row_num = 1
for r in regions:
    audit.base_client.set_region(r)
    audit_events = get_audit_events(audit, compartments, start_time, end_time)    
    if audit_events:
        for event in audit_events:
            if event.data.identity and event.data.identity.auth_type == 'natv': 
                user_groups = get_user_by_ocid(all_users,event.data.identity.principal_id)   
                formatted_event = format_audit_event(r, event, user_groups)
                event_id = (formatted_event["Region"], formatted_event["Compartment ID"], formatted_event["Event Name"], formatted_event["Principal ID"], formatted_event["Event Type"])
                
                if event_id not in unique_events:
                    unique_events.add(event_id)
                    for col, header in enumerate(headers):
                        ws.write(row_num, col, formatted_event[header])
                    row_num += 1
                    excel_file = r"C:\Security\Blogs\Identity\Access_Analyzer\logs\audit_events.xls"
                    wb.save(excel_file)


print(f"Audit events have been written to {excel_file}.")
now = datetime.datetime.now()
print("End time:", now)