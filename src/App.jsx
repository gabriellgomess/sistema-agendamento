import React, { useState } from 'react';
import { PageLayout } from './components/PageLayout';
import { loginRequest } from './authConfig';
import { ProfileData } from './components/ProfileData';
import AddEventForm from './components/AddEventForm';
import { callMsGraph, getCalendarEvents, updateCalendarEvent, deleteCalendarEvent } from './graph';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import './App.css';
import Button from 'react-bootstrap/Button';

/**
 * Renders information about the signed-in user or a button to retrieve data about the user
 */

const ProfileContent = () => {
    const { instance, accounts } = useMsal();
    const [graphData, setGraphData] = useState(null);
    const [calendarEvents, setCalendarEvents] = useState(null);
    const [editEvent, setEditEvent] = useState(null); // Para armazenar o evento que está sendo editado

    function RequestProfileData() {
        instance
            .acquireTokenSilent({
                ...loginRequest,
                account: accounts[0],
            })
            .then((response) => {
                callMsGraph(response.accessToken).then((response) => setGraphData(response));
            });
    }

    function RequestCalendarEvents() {
        instance
            .acquireTokenSilent({
                ...loginRequest,
                account: accounts[0],
            })
            .then((response) => {
                getCalendarEvents(response.accessToken).then((events) => setCalendarEvents(events.value));
            });
    }

    function handleDelete(eventId) {
        instance
            .acquireTokenSilent({
                ...loginRequest,
                account: accounts[0],
            })
            .then((response) => {
                deleteCalendarEvent(response.accessToken, eventId)
                    .then(() => RequestCalendarEvents()) // Recarrega a lista após deletar
                    .catch((error) => console.log(error));
            });
    }

    function handleEditSubmit(e) {
        e.preventDefault();
        instance
            .acquireTokenSilent({
                ...loginRequest,
                account: accounts[0],
            })
            .then((response) => {
                updateCalendarEvent(response.accessToken, editEvent.id, editEvent)
                    .then(() => {
                        setEditEvent(null); // Limpa o estado de edição
                        RequestCalendarEvents(); // Recarrega a lista após editar
                    })
                    .catch((error) => console.log(error));
            });
    }

    function handleEditChange(e) {
        const { name, value } = e.target;

        if (name.includes('start')) {
            setEditEvent((prevEvent) => ({
                ...prevEvent,
                start: {
                    ...prevEvent.start,
                    dateTime: new Date(value).toISOString(), // Simplesmente converte para ISO
                },
            }));
        } else if (name.includes('end')) {
            setEditEvent((prevEvent) => ({
                ...prevEvent,
                end: {
                    ...prevEvent.end,
                    dateTime: new Date(value).toISOString(), // Simplesmente converte para ISO
                },
            }));
        } else {
            setEditEvent((prevEvent) => ({
                ...prevEvent,
                [name]: value,
            }));
        }
    }




    return (
        <>
            <h5 className="profileContent">Olá {accounts[0].name}</h5>
            {graphData ? (
                <ProfileData graphData={graphData} />
            ) : (
                <Button variant="secondary" onClick={RequestProfileData}>
                    Request Profile
                </Button>
            )}

            {calendarEvents ? (
                <>
                    <h5>Eventos do Calendário:</h5>
                    <div style={{ display: 'flex', justifyContent: 'start' }}>
                        <ul style={{ textAlign: 'left' }}>
                            {calendarEvents.map((event) => (
                                <li key={event.id}>
                                    <strong>{event.subject}</strong> - {new Date(event.start.dateTime).toLocaleString()} to{' '}
                                    {new Date(event.end.dateTime).toLocaleString()}
                                    <Button onClick={() => setEditEvent(event)} variant="primary" style={{ marginLeft: '10px' }}>
                                        Edit
                                    </Button>
                                    <Button onClick={() => handleDelete(event.id)} variant="danger" style={{ marginLeft: '10px' }}>
                                        Delete
                                    </Button>
                                </li>
                            ))}
                        </ul>
                    </div>

                    {/* Formulário de edição de evento */}
                    {editEvent && (
                        <form onSubmit={handleEditSubmit}>
                            <h5>Edit Event</h5>
                            <div>
                                <label>Event Title:</label>
                                <input
                                    type="text"
                                    name="subject"
                                    value={editEvent.subject}
                                    onChange={handleEditChange}
                                    required
                                />
                            </div>
                            <div>
                                <label>Start Date and Time:</label>
                                <input
                                    type="datetime-local"
                                    name="start.dateTime"
                                    value={new Date(editEvent.start.dateTime).toISOString().slice(0, -1)} // Remove o "Z"
                                    onChange={handleEditChange}
                                    required
                                />
                            </div>
                            <div>
                                <label>End Date and Time:</label>
                                <input
                                    type="datetime-local"
                                    name="end.dateTime"
                                    value={new Date(editEvent.end.dateTime).toISOString().slice(0, -1)} // Remove o "Z"
                                    onChange={handleEditChange}
                                    required
                                />
                            </div>
                            <Button type="submit" variant="success">
                                Save Changes
                            </Button>
                        </form>
                    )}


                </>
            ) : (
                <Button variant="secondary" onClick={RequestCalendarEvents}>
                    Request Calendar Events
                </Button>
            )}
        </>
    );
};

/**
 * If a user is authenticated the ProfileContent component above is rendered. Otherwise a message indicating a user is not authenticated is rendered.
 */
const MainContent = () => {
    return (
        <div className="App">
            <AuthenticatedTemplate>
                <ProfileContent />
                <AddEventForm onEventAdded={() => RequestCalendarEvents()} />
            </AuthenticatedTemplate>

            <UnauthenticatedTemplate>
                <h5 className="card-title">Please sign-in to see your profile information.</h5>
            </UnauthenticatedTemplate>
        </div>
    );
};

export default function App() {
    return (
        <PageLayout>
            <MainContent />
        </PageLayout>
    );
}
