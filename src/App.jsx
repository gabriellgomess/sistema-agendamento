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

const ProfileContent = ({ RequestCalendarEvents, calendarEvents, setCalendarEvents }) => {
    const { instance, accounts } = useMsal();
    const [graphData, setGraphData] = useState(null);
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
    
        // Aqui, garantimos que o fuso horário seja incluído nos detalhes do evento
        const updatedEvent = {
            ...editEvent,
            start: {
                ...editEvent.start,
                timeZone: "America/Sao_Paulo" // Informar explicitamente o fuso horário correto
            },
            end: {
                ...editEvent.end,
                timeZone: "America/Sao_Paulo" // Informar explicitamente o fuso horário correto
            }
        };
    
        instance
            .acquireTokenSilent({
                ...loginRequest,
                account: accounts[0],
            })
            .then((response) => {
                updateCalendarEvent(response.accessToken, editEvent.id, updatedEvent)
                    .then(() => {
                        setEditEvent(null); // Limpa o estado de edição
                        RequestCalendarEvents(); // Recarrega a lista após editar
                    })
                    .catch((error) => console.log(error));
            });
    }
    
    

    function handleEditChange(e) {
        const { name, value } = e.target;
    
        // Aqui convertendo a data para UTC usando toISOString
        const dateInUTC = new Date(value).toISOString();
    
        if (name.includes('start')) {
            setEditEvent((prevEvent) => ({
                ...prevEvent,
                start: {
                    ...prevEvent.start,
                    dateTime: dateInUTC // Converte para UTC
                }
            }));
        } else if (name.includes('end')) {
            setEditEvent((prevEvent) => ({
                ...prevEvent,
                end: {
                    ...prevEvent.end,
                    dateTime: dateInUTC // Converte para UTC
                }
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
                    <div style={{ display: 'flex', justifyContent: 'start', gap: '15px', flexWrap: 'wrap'}}>
                        
                            {calendarEvents.map((event) => (
                                <div key={event.id} style={{border: '1px solid grey', borderRadius: '10px', width: '400px', padding: '15px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between'  }}>
                                    <h6>{event.subject}</h6>
                                    <p>{new Date(event.start.dateTime).toLocaleString()} as{' '}
                                    {new Date(event.end.dateTime).toLocaleString()}</p>
                                    <div style={{display: 'flex', justifyContent: 'center'}}>
                                      <Button onClick={() => setEditEvent(event)} variant="primary" style={{ marginLeft: '10px' }}>
                                        Editar
                                    </Button>
                                    <Button onClick={() => handleDelete(event.id)} variant="danger" style={{ marginLeft: '10px' }}>
                                        Deletar
                                    </Button>  
                                    </div>
                                    
                                </div>
                            ))}
                        
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
                                    value={new Date(editEvent.start.dateTime).toLocaleString('sv-SE', { timeZone: 'America/Sao_Paulo' }).replace(' ', 'T')}
                                    onChange={handleEditChange}
                                    required
                                />
                            </div>
                            <div>
                                <label>End Date and Time:</label>
                                <input
                                    type="datetime-local"
                                    name="end.dateTime"
                                    value={new Date(editEvent.end.dateTime).toLocaleString('sv-SE', { timeZone: 'America/Sao_Paulo' }).replace(' ', 'T')}
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
    const { instance, accounts } = useMsal();
    const [calendarEvents, setCalendarEvents] = useState(null); // Mover o estado para MainContent

    const RequestCalendarEvents = () => {
        instance
            .acquireTokenSilent({
                ...loginRequest,
                account: accounts[0],
            })
            .then((response) => {
                getCalendarEvents(response.accessToken).then((events) => setCalendarEvents(events.value)); // setCalendarEvents agora é acessível aqui
            });
    };

    return (
        <div className="App">
            <AuthenticatedTemplate>
                <ProfileContent
                    RequestCalendarEvents={RequestCalendarEvents}
                    calendarEvents={calendarEvents}
                    setCalendarEvents={setCalendarEvents} // Passando como props
                />
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
