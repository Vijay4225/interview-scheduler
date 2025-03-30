import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO

def is_available(person, start, end):
    """Check if a time slot is free for a person."""
    for booked_start, booked_end in person['booked_slots']:
        if not (end <= booked_start or start >= booked_end):
            return False
    return True

def main():
    st.title("Automated Interview Scheduler 🗓️")
    
    # Upload Excel files
    interviewers_file = st.file_uploader("Upload Interviewers Excel", type=["xlsx"])
    interviewees_file = st.file_uploader("Upload Interviewees Excel", type=["xlsx"])
    
    if interviewers_file and interviewees_file:
        # Load data
        interviewers_df = pd.read_excel(interviewers_file)
        interviewees_df = pd.read_excel(interviewees_file)
        
        # Preprocess data
        interviewers = []
        for _, row in interviewers_df.iterrows():
            interviewers.append({
                "id": row["ID"],
                "name": row["Name"],
                "skills": [s.strip() for s in str(row["Skills"]).split(",")],
                "available_slots": [(pd.to_datetime(row["Available_Start"]), pd.to_datetime(row["Available_End"]))],
                "booked_slots": []
            })
        
        interviewees = []
        for _, row in interviewees_df.iterrows():
            interviewees.append({
                "id": row["ID"],
                "name": row["Name"],
                "required_skill": row["Required_Skill"],
                "duration": row["Duration"],
                "available_slots": [(pd.to_datetime(row["Available_Start"]), pd.to_datetime(row["Available_End"]))],
                "booked_slots": []
            })
        
        # Schedule interviews
        schedule = []
        for interviewee in interviewees:
            req_skill = interviewee["required_skill"]
            req_duration = interviewee["duration"]
            
            for avl_start, avl_end in interviewee["available_slots"]:
                # Check if slot has enough time
                if (avl_end - avl_start).total_seconds() / 60 < req_duration:
                    continue
                
                # Find eligible interviewers
                eligible = [i for i in interviewers if req_skill in i["skills"]]
                for interviewer in eligible:
                    for i_avl_start, i_avl_end in interviewer["available_slots"]:
                        # Find overlapping window
                        overlap_start = max(avl_start, i_avl_start)
                        overlap_end = min(avl_end, i_avl_end)
                        if overlap_start >= overlap_end:
                            continue
                        
                        # Check overlap duration
                        overlap_min = (overlap_end - overlap_start).total_seconds() / 60
                        if overlap_min < req_duration:
                            continue
                        
                        # Check time slots in 15-minute increments
                        current_start = overlap_start
                        while current_start + timedelta(minutes=req_duration) <= overlap_end:
                            current_end = current_start + timedelta(minutes=req_duration)
                            if is_available(interviewee, current_start, current_end) and \
                               is_available(interviewer, current_start, current_end):
                                # Schedule the interview
                                schedule.append({
                                    "Interviewee": interviewee["name"],
                                    "Interviewer": interviewer["name"],
                                    "Skill": req_skill,
                                    "Start": current_start.strftime("%Y-%m-%d %H:%M"),
                                    "End": current_end.strftime("%Y-%m-%d %H:%M"),
                                    "Duration (mins)": req_duration
                                })
                                # Block the time
                                interviewee["booked_slots"].append((current_start, current_end))
                                interviewer["booked_slots"].append((current_start, current_end))
                                break  # Move to next interviewee
                            current_start += timedelta(minutes=15)
                        else:
                            continue  # No slot found in this overlap
                        break  # Slot found, move to next interviewee
                    else:
                        continue  # No interviewer slot matched
                    break  # Interviewer matched
                else:
                    continue  # No eligible interviewers
                break  # Slot found
        
        # Display and export results
        if schedule:
            st.success(f"Scheduled {len(schedule)} interviews!")
            schedule_df = pd.DataFrame(schedule)
            st.dataframe(schedule_df)
            
            # Export to Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                schedule_df.to_excel(writer, index=False)
            st.download_button(
                label="Download Schedule",
                data=output.getvalue(),
                file_name="Interview_Schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No interviews scheduled. Check input data for conflicts.")

if __name__ == "__main__":
    main()