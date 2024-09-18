import React, { useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { createCalendarEvent } from '../graph';
import { loginRequest } from '../authConfig';

const AddEventForm = () => {
    const { instance, accounts } = useMsal();
    const [eventDetails, setEventDetails] = useState({
        subject: "",
        start: {
            dateTime: "",
            timeZone: "America/Sao_Paulo"
        },
        end: {
            dateTime: "",
            timeZone: "America/Sao_Paulo"
        }
    });

    const handleInputChange = (e) => {
        const { name, value } = e.target;
        setEventDetails(prevDetails => ({
            ...prevDetails,
            [name]: value
        }));
    };

    const handleDateTimeChange = (e, field) => {
        const { value } = e.target;

        // Converte para ISO 8601
        const isoDateTime = new Date(value).toISOString();

        setEventDetails(prevDetails => ({
            ...prevDetails,
            [field]: {
                ...prevDetails[field],
                dateTime: isoDateTime // Define a data como ISO
            }
        }));
    };

    const handleSubmit = (e) => {
        e.preventDefault();
        instance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0],
        }).then((response) => {
            createCalendarEvent(response.accessToken, eventDetails)
                .then(response => {
                    console.log("Event created successfully:", response);
                })
                .catch(error => console.log(error));
        });
    };

    return (
        <form onSubmit={handleSubmit}>
            <div>
                <label>Event Title:</label>
                <input type="text" name="subject" value={eventDetails.subject} onChange={handleInputChange} required />
            </div>
            <div>
                <label>Start Date and Time:</label>
                <input type="datetime-local" name="dateTime" value={eventDetails.start.dateTime.slice(0, 16)} onChange={(e) => handleDateTimeChange(e, "start")} required />
            </div>
            <div>
                <label>End Date and Time:</label>
                <input type="datetime-local" name="dateTime" value={eventDetails.end.dateTime.slice(0, 16)} onChange={(e) => handleDateTimeChange(e, "end")} required />
            </div>
            <button type="submit">Add Event</button>
        </form>
    );
};

export default AddEventForm;
