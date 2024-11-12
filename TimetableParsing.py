import pandas as pd
import json

excel_file_path = 'C:\\PythonProjects\\Timetable Workbook - SUTT Task 1.xlsx'  # File path replaced
df = pd.read_excel(excel_file_path, sheet_name=None, header=None)

parsed_data = {}

def clean_value(value):
    return value if pd.notna(value) else None

def map_day_to_letter(day):
    day_map = {
        "Mon": "M",
        "Tue": "T",
        "Wed": "W",
        "Thu": "Th",
        "Fri": "F"
    }
    return day_map.get(day, day)  

def convert_time_to_hour(time):
    try:
        return int(time) if pd.notna(time) else None
    except ValueError:
        return None

for sheet_name, sheet_data in df.items():
    sheet_data = sheet_data.iloc[3:]

    sheet_parsed = []
    current_course_title = None
    current_course_number = None
    last_section_data = None  

    for _, row in sheet_data.iterrows():
        course_code = clean_value(row[1])
        course_title = clean_value(row[2])
        section = clean_value(row[6]) 
        instructor = clean_value(row[7])  
        time_slot = clean_value(row[9])  
        room = clean_value(row[8])  
        midsem = clean_value(row[10]) 
        compre = clean_value(row[11]) 

        if pd.notna(course_title):
            current_course_title = course_title
            current_course_number = course_code

        if current_course_title is None:
            continue  

        course_data = next((course for course in sheet_parsed if course['course_code'] == current_course_number), None)

        if not course_data:
            course_data = {
                'course_code': current_course_number,
                'course_title': current_course_title,
                'credit_structure': {
                    'lecture': clean_value(row[3]), 
                    'practical': clean_value(row[4]),
                    'unit': clean_value(row[5]) 
                },
                'sections': []
            }
            sheet_parsed.append(course_data) 

        if pd.notna(section):
            section_data = next((sec for sec in course_data['sections'] if sec['section'] == section), None)

            if not section_data:
                section_data = {
                    'section': section,  
                    'instructors': [], 
                    'room': room,  
                    'timing': []  
                }
                course_data['sections'].append(section_data)

            last_section_data = section_data  

        if pd.notna(instructor) and last_section_data:
            last_section_data['instructors'].append(instructor)

            # Process the time slot (convert to days and hours)
            if pd.notna(time_slot):
                time_parts = time_slot.split()
                days = time_parts[0].split('-')
                slots = [convert_time_to_hour(hour) for hour in time_parts[1].split('-')]

                for day in days:
                    last_section_data['timing'].append({
                        'day': map_day_to_letter(day),  
                        'slots': slots  
                    })

        if pd.isna(midsem) and pd.isna(compre):
            continue
        if pd.notna(midsem):
            last_section_data['midsem'] = midsem
        if pd.notna(compre):
            last_section_data['compre'] = compre

    parsed_data[sheet_name] = sheet_parsed

output_json = json.dumps(parsed_data, indent=4)

with open('parsed_data.json', 'w') as json_file:
    json_file.write(output_json)

print("Parsing completed and saved to parsed_data.json")