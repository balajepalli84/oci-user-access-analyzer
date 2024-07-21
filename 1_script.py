import datetime
import oci
import xlwt

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
        list_events_response = oci.pagination.list_call_get_all_results(
            audit.list_events,
            compartment_id=c,
            start_time=start_time,
            end_time=end_time).data

        list_of_audit_events.extend(list_events_response)
    return list_of_audit_events

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
        "User Groups": user_groups
    }

def user_group_info(user_ocid):
    group_names = []
    list_user_group_memberships_response = identity.list_user_group_memberships(
        compartment_id=config["tenancy"],
        user_id=user_ocid
    )
    for e in list_user_group_memberships_response.data:
        get_group_response = identity.get_group(group_id=e.group_id)
        group_names.append(get_group_response.data.name)

    return ', '.join(group_names)


config = oci.config.from_file()
tenancy_id = config["tenancy"]

identity = oci.identity.IdentityClient(config)
end_time = datetime.datetime.utcnow()
start_time = end_time + datetime.timedelta(days=-5)
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
row_num = 1
for r in regions:
    audit.base_client.set_region(r)
    audit_events = get_audit_events(audit, compartments, start_time, end_time)    
    if audit_events:
        for event in audit_events:
            if event.data.identity and event.data.identity.auth_type == 'natv':
                user_groups = user_group_info(event.data.identity.principal_id)            
                formatted_event = format_audit_event(r, event, user_groups)
                event_id = (formatted_event["Region"], formatted_event["Compartment ID"], formatted_event["Event Name"], formatted_event["Principal ID"], formatted_event["Event Type"])
                
                if event_id not in unique_events:
                    unique_events.add(event_id)
                    for col, header in enumerate(headers):
                        ws.write(row_num, col, formatted_event[header])
                    row_num += 1
                    excel_file = r"C:\Security\Blogs\Access_Analyzer\audit_events.xls"
                    wb.save(excel_file)

print(f"Audit events have been written to {excel_file}.")
