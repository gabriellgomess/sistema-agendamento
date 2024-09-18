import { graphConfig } from "./authConfig";

/**
 * Attaches a given access token to a MS Graph API call. Returns information about the user
 * @param accessToken 
 */
export async function callMsGraph(accessToken) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch(graphConfig.graphMeEndpoint, options)
        .then(response => response.json())
        .catch(error => console.log(error));
}

/**
 * Fetches the list of calendar events from the signed-in user's calendar.
 * @param accessToken 
 */
export async function getCalendarEvents(accessToken) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    return fetch("https://graph.microsoft.com/v1.0/me/events", options)
        .then(response => response.json())
        .catch(error => console.log(error));
}

/**
 * Creates a new event in the signed-in user's calendar.
 * @param accessToken
 * @param eventDetails
 */
export async function createCalendarEvent(accessToken, eventDetails) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);
    headers.append("Content-Type", "application/json");

        // Adicionar timeZone nos detalhes do evento
        eventDetails.start.timeZone = "America/Sao_Paulo"; 
        eventDetails.end.timeZone = "America/Sao_Paulo";

    const options = {
        method: "POST",
        headers: headers,
        body: JSON.stringify(eventDetails),
    };

    return fetch("https://graph.microsoft.com/v1.0/me/events", options)
        .then(response => response.json())
        .catch(error => console.log(error));
}

/**
 * Updates an existing event in the user's calendar.
 * @param accessToken
 * @param eventId
 * @param updatedEventDetails
 */
export async function updateCalendarEvent(accessToken, eventId, updatedEventDetails) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);
    headers.append("Content-Type", "application/json");

    // Adicionar timeZone nos detalhes do evento
    updatedEventDetails.start.timeZone = "America/Sao_Paulo";
    updatedEventDetails.end.timeZone = "America/Sao_Paulo";

    const options = {
        method: "PATCH",
        headers: headers,
        body: JSON.stringify(updatedEventDetails), // Aqui os detalhes são enviados com UTC e timeZone correto
    };

    return fetch(`https://graph.microsoft.com/v1.0/me/events/${eventId}`, options)
        .then(response => response.json())
        .catch(error => console.log(error));
}



/**
 * Deletes an event from the user's calendar.
 * @param accessToken
 * @param eventId
 */
export async function deleteCalendarEvent(accessToken, eventId) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "DELETE",
        headers: headers
    };

    return fetch(`https://graph.microsoft.com/v1.0/me/events/${eventId}`, options)
        .then(response => {
            if (response.status === 204) {
                // Sucesso ao deletar
                return "Deleted";
            } else {
                throw new Error("Error deleting event");
            }
        })
        .catch(error => console.log(error));
}



