import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import timedelta
from io import BytesIO
import base64
from streamlit.components.v1 import html

def is_available(person, start, end):
    for booked_start, booked_end in person['booked_slots']:
        if not (end <= booked_start or start >= booked_end):
            return False
    return True

def create_gantt_chart(df, filter_type, filter_value):
    """Create interactive Gantt chart with filters"""
    if filter_value != "All":
        if filter_type == "skill":
            df = df[df["Skill"] == filter_value]
        else:
            df = df[df["Interviewer"] == filter_value]
    
    fig = px.timeline(
        df,
        x_start="Start",
        x_end="End",
        y="Interviewer",
        color="Skill",
        hover_name="Interviewee",
        title=f"Interview Schedule - {filter_value if filter_value != 'All' else 'All Interviews'}",
        labels={"Interviewer": "Interviewer/Skill"},
        height=600
    )
    fig.update_yaxes(categoryorder="total ascending")
    fig.update_layout(
        xaxis_title="Timeline",
        yaxis_title="Interviewer" if filter_type == "interviewer" else "Skill",
        hovermode="closest"
    )
    return fig

def main():
    st.title("📅 Interactive Interview Scheduler")
    
    # File upload section
    col1, col2 = st.columns(2)
    with col1:
        interviewers_file = st.file_uploader("Upload Interviewers Excel", type=["xlsx"])
    with col2:
        interviewees_file = st.file_uploader("Upload Interviewees Excel", type=["xlsx"])
    
    if interviewers_file and interviewees_file:
        # ... [Keep existing data loading and scheduling logic] ...
        
        # After generating schedule and unscheduled lists
        # ===================================================
        # ADD MISSING DATA LOADING LOGIC
        # ===================================================
        interviewers_df = pd.read_excel(interviewers_file)
        interviewees_df = pd.read_excel(interviewees_file)

        # Preprocess interviewers
        interviewers = []
        for _, row in interviewers_df.iterrows():
            interviewers.append({
                "id": row["ID"],
                "name": row["Name"],
                "skills": [s.strip() for s in str(row["Skills"]).split(",")],
                "available_slots": [(pd.to_datetime(row["Available_Start"]), pd.to_datetime(row["Available_End"]))],
                "booked_slots": []
            })

        # Preprocess interviewees
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
        
        # Schedule interviews (your existing logic)
        schedule = []
        unscheduled = []
        
        for interviewee in interviewees:
            scheduled = False
            req_skill = interviewee["required_skill"]
            req_duration = interviewee["duration"]
            
            for avl_start, avl_end in interviewee["available_slots"]:
                if (avl_end - avl_start).total_seconds() / 60 < req_duration:
                    continue
                
                eligible_interviewers = [i for i in interviewers if req_skill in i["skills"]]
                if not eligible_interviewers:  # NEW: Check for no eligible interviewers
                    unscheduled.append({
                        **interviewee,
                        "reason": "No matching interviewer"
                    })
                    break
                
                for interviewer in eligible_interviewers:
                    for i_avl_start, i_avl_end in interviewer["available_slots"]:
                        overlap_start = max(avl_start, i_avl_start)
                        overlap_end = min(avl_end, i_avl_end)
                        if overlap_start >= overlap_end:
                            continue
                        
                        overlap_min = (overlap_end - overlap_start).total_seconds() / 60
                        if overlap_min < req_duration:
                            continue
                        
                        current_start = overlap_start
                        while current_start + timedelta(minutes=req_duration) <= overlap_end:
                            current_end = current_start + timedelta(minutes=req_duration)
                            if is_available(interviewee, current_start, current_end) and \
                               is_available(interviewer, current_start, current_end):
                                schedule.append({
                                    "Interviewee": interviewee["name"],
                                    "Interviewer": interviewer["name"],
                                    "Skill": req_skill,
                                    "Start": current_start.strftime("%Y-%m-%d %H:%M"),
                                    "End": current_end.strftime("%Y-%m-%d %H:%M"),
                                    "Duration (mins)": req_duration
                                })
                                interviewee["booked_slots"].append((current_start, current_end))
                                interviewer["booked_slots"].append((current_start, current_end))
                                scheduled = True
                                break
                            current_start += timedelta(minutes=15)
                        if scheduled:
                            break
                    if scheduled:
                        break
                if scheduled:
                    break
            if not scheduled:  # NEW: Add to unscheduled list
                unscheduled.append({
                    "ID": interviewee["id"],
                    "Name": interviewee["name"],
                    "Required Skill": req_skill,
                    "Duration": req_duration,
                    "Available Start": interviewee["available_slots"][0][0].strftime("%Y-%m-%d %H:%M"),
                    "Available End": interviewee["available_slots"][0][1].strftime("%Y-%m-%d %H:%M"),
                    "Reason": "No available slots"
                })
        # ... [Keep your scheduling logic here] ...

        # Generate Excel output
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            scheduled_df = pd.DataFrame(schedule)
            scheduled_df.to_excel(writer, index=False, sheet_name='Scheduled')
            
            unscheduled_df = pd.DataFrame(unscheduled)
            unscheduled_df.to_excel(writer, index=False, sheet_name='Unscheduled')
        scheduled_df = pd.DataFrame(schedule)
        
        # Visualization Section
        st.markdown("---")
        st.header("📊 Interactive Schedule Visualization")
        
        # Create filters
        col1, col2 = st.columns(2)
        with col1:
            filter_type = st.radio("Filter by:", ["skill", "interviewer"])
        with col2:
            if filter_type == "skill":
                skills = sorted(scheduled_df["Skill"].unique())
                selected_filter = st.selectbox("Select Skill:", ["All"] + skills)
            else:
                interviewers = sorted(scheduled_df["Interviewer"].unique())
                selected_filter = st.selectbox("Select Interviewer:", ["All"] + interviewers)
        
        # Create and display Gantt chart
        if not scheduled_df.empty:
            fig = create_gantt_chart(scheduled_df, filter_type, selected_filter)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("No interviews scheduled to visualize")
        
        # Auto-download and backup button (keep existing functionality)
        # ... [Keep existing download logic] ...
        excel_data = output.getvalue()
        b64 = base64.b64encode(excel_data).decode()
        html(f'<script>var a=document.createElement("a");a.href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}";a.download="Interview_Report.xlsx";document.body.appendChild(a);a.click();</script>', height=0)
        
        st.download_button(
            "📥 Download Report Again", 
            data=excel_data,
            file_name="Interview_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
