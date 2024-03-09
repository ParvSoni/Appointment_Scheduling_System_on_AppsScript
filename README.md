# Appointment-Scheduling-System-on-AppsScript
### Client Submission Form:

Clients fill out a form with their details, including Name, Reason for Appointment, Date, Time, and Email address.
This information is collected through a Google Form, which automatically populates a Google Sheets spreadsheet with the responses.
### Email Notifications:

After submitting the form, the client receives an email confirmation ("Request Received") on the provided Gmail address.
Simultaneously, the doctor is notified via email of a new appointment request ("New request received"). This email contains a link to the Google Sheets file that stores all the responses.
### Doctor's Approval/Rejection:

The doctor reviews the appointment requests in the Google Sheets file.
The doctor can either approve or reject an appointment request.
Upon making a decision, the client receives either a confirmation email or a rescheduling email, depending on the doctor's decision.
### Appointment Reminders:

If the doctor approves an appointment, a reminder is set on the doctor's personal calendar 30 minutes before the scheduled appointment time.
The client also receives a new confirmation email when the appointment is approved.
### Rescheduling Functionality:

If the doctor rejects an appointment request, the client receives a rescheduling email.
The email contains a button that directs the client to the appointment form for resubmission.
### Conflict Detection:

The system automatically checks for scheduling conflicts.
If a new client submits a form for an appointment on a date and time that's already booked, the system sends a conflicting email to the client.
The email includes a reschedule button that allows the client to select an alternative appointment slot.
### Customization:

The system allows for customization of the doctor's personal Gmail address and the duration of appointment slots, making it adaptable to different professional's needs.

# Video Description of the Project 
[recording-2023-05-18-11-21-16.webm.webm](https://github.com/ParvSoni/Appointment-Scheduling-System-on-AppsScript/assets/123165567/e855a4fe-afc4-4262-bfc1-9baa81e205f8)
