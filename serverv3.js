#!/usr/bin/env node

/**
 * Google Workspace MCP Server
 * Integrates Gmail, Google Calendar, and Google Docs with Model Context Protocol
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
    CallToolRequestSchema,
    ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { google } from "googleapis";
import { authenticate } from "@google-cloud/local-auth";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// RUTAS CORRECTAS - Se guardan en la carpeta del proyecto
const TOKEN_PATH = path.join(__dirname, "token.json");
const CREDENTIALS_PATH = path.join(__dirname, "credentials.json");

// Scopes for Gmail, Calendar, and Docs
const SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive.readonly",
    // 🆕 SCOPES DE GOOGLE FIT
    "https://www.googleapis.com/auth/fitness.activity.read", // Para leer datos agregados
    "https://www.googleapis.com/auth/fitness.activity.write", // Para registrar sesiones
];

class GoogleWorkspaceMCPServer {
    constructor() {
        this.server = new Server(
            {
                name: "google-workspace-mcp",
                version: "1.0.0",
            },
            {
                capabilities: {
                    tools: {},
                },
            }
        );

        this.auth = null;
        this.gmail = null;
        this.calendar = null;
        this.docs = null;
        this.drive = null;

        this.setupHandlers();
    }

    async loadSavedCredentialsIfExist() {
        try {
            let content;
            let keys;
            
            // Intentar cargar desde variables de entorno primero
            if (process.env.GOOGLE_TOKEN) {
                content = process.env.GOOGLE_TOKEN;
                const credentials = JSON.parse(content);
                
                keys = process.env.GOOGLE_CREDENTIALS 
                    ? JSON.parse(process.env.GOOGLE_CREDENTIALS)
                    : JSON.parse(await fs.readFile(CREDENTIALS_PATH));
                    
                const key = keys.installed || keys.web;
                const client = new google.auth.OAuth2(
                    key.client_id,
                    key.client_secret,
                    key.redirect_uris[0]
                );

                client.setCredentials({
                    refresh_token: credentials.refresh_token,
                    access_token: credentials.access_token,
                    token_type: credentials.token_type,
                    expiry_date: credentials.expiry_date,
                });

                console.error("✅ Token cargado desde variables de entorno");
                return client;
            }
            
            // Si no hay variables de entorno, cargar desde archivos
            content = await fs.readFile(TOKEN_PATH);
            const credentials = JSON.parse(content);

            // Leer las credenciales originales para obtener client_id y client_secret
            keys = JSON.parse(await fs.readFile(CREDENTIALS_PATH));
            const key = keys.installed || keys.web;

            // Crear OAuth2Client con las credenciales correctas
            const client = new google.auth.OAuth2(
                key.client_id,
                key.client_secret,
                key.redirect_uris[0]
            );

            // Establecer las credenciales guardadas
            client.setCredentials({
                refresh_token: credentials.refresh_token,
                access_token: credentials.access_token,
                token_type: credentials.token_type,
                expiry_date: credentials.expiry_date,
            });

            console.error("✅ Token cargado desde archivo");
            return client;
        } catch (err) {
            console.error("ℹ️ No se encontró token guardado, se requiere autenticación");
            return null;
        }
    }

    async saveCredentials(client) {
        const content = await fs.readFile(CREDENTIALS_PATH);
        const keys = JSON.parse(content);
        const key = keys.installed || keys.web;

        const payload = JSON.stringify({
            type: "authorized_user",
            client_id: key.client_id,
            client_secret: key.client_secret,
            refresh_token: client.credentials.refresh_token,
            access_token: client.credentials.access_token,
            token_type: client.credentials.token_type,
            expiry_date: client.credentials.expiry_date,
        });

        await fs.writeFile(TOKEN_PATH, payload);
        console.error("✅ Credenciales guardadas exitosamente");
    }

    async ensureValidToken() {
        if (!this.auth || !this.auth.credentials) {
            throw new Error("No hay autenticación disponible");
        }

        const now = new Date().getTime();
        const expiry = this.auth.credentials.expiry_date;

        // Si no hay expiry_date o el token expiró, renovarlo
        if (!expiry || expiry <= now) {
            console.error("🔄 Token expirado, renovando...");
            try {
                const { credentials } = await this.auth.refreshAccessToken();
                this.auth.setCredentials(credentials);
                await this.saveCredentials(this.auth);
                console.error("✅ Token renovado exitosamente");
            } catch (error) {
                console.error("❌ Error al renovar token:", error.message);
                console.error("⚠️ Necesitas volver a autenticarte. Borra token.json y reinicia.");
                throw new Error("Token inválido, se requiere re-autenticación");
            }
        } else if (expiry - now < 5 * 60 * 1000) {
            // Si expira en menos de 5 minutos, renovarlo preventivamente
            console.error("🔄 Token por expirar, renovando preventivamente...");
            try {
                const { credentials } = await this.auth.refreshAccessToken();
                this.auth.setCredentials(credentials);
                await this.saveCredentials(this.auth);
                console.error("✅ Token renovado preventivamente");
            } catch (error) {
                console.error("⚠️ Error al renovar token preventivamente:", error.message);
            }
        }
    }

    async authorize() {
        let client = await this.loadSavedCredentialsIfExist();
        if (client) {
            // Verificar que el token sea válido
            try {
                await client.getAccessToken();
                return client;
            } catch (error) {
                console.error("⚠️ Token guardado inválido, requiere re-autenticación");
                client = null;
            }
        }

        // Si no hay cliente válido, autenticar desde cero
        client = await authenticate({
            scopes: SCOPES,
            keyfilePath: CREDENTIALS_PATH,
        });

        if (client.credentials) {
            await this.saveCredentials(client);
        }

        return client;
    }

    async initialize() {
        try {
            this.auth = await this.authorize();
            this.gmail = google.gmail({ version: "v1", auth: this.auth });
            this.calendar = google.calendar({ version: "v3", auth: this.auth });
            this.docs = google.docs({ version: "v1", auth: this.auth });
            this.drive = google.drive({ version: "v3", auth: this.auth });
            // 🆕 INICIALIZACIÓN DE GOOGLE FIT
            this.fit = google.fitness({ version: "v1", auth: this.auth });
            console.error("✅ Autenticación exitosa - Servidor listo");
        } catch (error) {
            console.error("❌ Error durante la autenticación:", error.message);
            throw error;
        }
    }

    setupHandlers() {
        this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
            tools: [
                // Gmail Tools
                {
                    name: "gmail_list_messages",
                    description: "Lista tus emails recientes con filtros opcionales",
                    inputSchema: {
                        type: "object",
                        properties: {
                            query: {
                                type: "string",
                                description: "Búsqueda de Gmail (ej: 'is:unread', 'from:user@example.com')",
                            },
                            maxResults: {
                                type: "number",
                                description: "Máximo de mensajes a retornar (default: 10)",
                                default: 10,
                            },
                        },
                    },
                },
                {
                    name: "gmail_get_message",
                    description: "Obtiene un email específico por su ID",
                    inputSchema: {
                        type: "object",
                        properties: {
                            messageId: {
                                type: "string",
                                description: "ID del mensaje a obtener",
                            },
                        },
                        required: ["messageId"],
                    },
                },
                {
                    name: "gmail_send_message",
                    description: "Envía un email a través de Gmail",
                    inputSchema: {
                        type: "object",
                        properties: {
                            to: {
                                type: "string",
                                description: "Dirección de email del destinatario",
                            },
                            subject: {
                                type: "string",
                                description: "Asunto del email",
                            },
                            body: {
                                type: "string",
                                description: "Cuerpo del email (texto plano)",
                            },
                        },
                        required: ["to", "subject", "body"],
                    },
                },
                {
                    name: "gmail_create_draft",
                    description: "Crea un borrador de email en Gmail",
                    inputSchema: {
                        type: "object",
                        properties: {
                            to: {
                                type: "string",
                                description: "Dirección de email del destinatario",
                            },
                            subject: {
                                type: "string",
                                description: "Asunto del email",
                            },
                            body: {
                                type: "string",
                                description: "Cuerpo del email (texto plano)",
                            },
                        },
                        required: ["to", "subject", "body"],
                    },
                },
                // Calendar Tools
                {
                    name: "calendar_list_events",
                    description: "Lista los próximos eventos del calendario",
                    inputSchema: {
                        type: "object",
                        properties: {
                            maxResults: {
                                type: "number",
                                description: "Máximo de eventos a retornar (default: 10)",
                                default: 10,
                            },
                            timeMin: {
                                type: "string",
                                description: "Fecha/hora mínima para eventos (ISO 8601, default: ahora)",
                            },
                            timeMax: {
                                type: "string",
                                description: "Fecha/hora máxima para eventos (ISO 8601)",
                            },
                        },
                    },
                },
                {
                    name: "calendar_create_event",
                    description: "Crea un nuevo evento en el calendario",
                    inputSchema: {
                        type: "object",
                        properties: {
                            summary: {
                                type: "string",
                                description: "Título del evento",
                            },
                            description: {
                                type: "string",
                                description: "Descripción del evento",
                            },
                            startDateTime: {
                                type: "string",
                                description: "Fecha y hora de inicio (formato ISO 8601)",
                            },
                            endDateTime: {
                                type: "string",
                                description: "Fecha y hora de fin (formato ISO 8601)",
                            },
                            location: {
                                type: "string",
                                description: "Ubicación del evento",
                            },
                            attendees: {
                                type: "array",
                                items: { type: "string" },
                                description: "Lista de emails de asistentes",
                            },
                        },
                        required: ["summary", "startDateTime", "endDateTime"],
                    },
                },
                {
                    name: "calendar_update_event",
                    description: "Actualiza un evento existente del calendario",
                    inputSchema: {
                        type: "object",
                        properties: {
                            eventId: {
                                type: "string",
                                description: "ID del evento a actualizar",
                            },
                            summary: {
                                type: "string",
                                description: "Título del evento",
                            },
                            description: {
                                type: "string",
                                description: "Descripción del evento",
                            },
                            startDateTime: {
                                type: "string",
                                description: "Fecha y hora de inicio (formato ISO 8601)",
                            },
                            endDateTime: {
                                type: "string",
                                description: "Fecha y hora de fin (formato ISO 8601)",
                            },
                            location: {
                                type: "string",
                                description: "Ubicación del evento",
                            },
                        },
                        required: ["eventId"],
                    },
                },
                {
                    name: "calendar_delete_event",
                    description: "Elimina un evento del calendario",
                    inputSchema: {
                        type: "object",
                        properties: {
                            eventId: {
                                type: "string",
                                description: "ID del evento a eliminar",
                            },
                        },
                        required: ["eventId"],
                    },
                },
                {
                    name: "calendar_find_free_slots",
                    description: "Encuentra espacios libres en el calendario",
                    inputSchema: {
                        type: "object",
                        properties: {
                            startDate: {
                                type: "string",
                                description: "Fecha de inicio para buscar (formato ISO 8601)",
                            },
                            endDate: {
                                type: "string",
                                description: "Fecha de fin para buscar (formato ISO 8601)",
                            },
                            duration: {
                                type: "number",
                                description: "Duración deseada en minutos",
                            },
                        },
                        required: ["startDate", "endDate", "duration"],
                    },
                },
                // Google Docs Tools
                {
                    name: "docs_create",
                    description: "Crea un nuevo documento de Google Docs",
                    inputSchema: {
                        type: "object",
                        properties: {
                            title: {
                                type: "string",
                                description: "Título del documento",
                            },
                            content: {
                                type: "string",
                                description: "Contenido inicial del documento (texto plano)",
                            },
                        },
                        required: ["title"],
                    },
                },
                {
                    name: "docs_get",
                    description: "Obtiene el contenido de un documento de Google Docs",
                    inputSchema: {
                        type: "object",
                        properties: {
                            documentId: {
                                type: "string",
                                description: "ID del documento a obtener",
                            },
                        },
                        required: ["documentId"],
                    },
                },
                {
                    name: "docs_append_text",
                    description: "Añade texto al final de un documento",
                    inputSchema: {
                        type: "object",
                        properties: {
                            documentId: {
                                type: "string",
                                description: "ID del documento",
                            },
                            text: {
                                type: "string",
                                description: "Texto a añadir",
                            },
                        },
                        required: ["documentId", "text"],
                    },
                },
                {
                    name: "docs_insert_text",
                    description: "Inserta texto en una posición específica del documento",
                    inputSchema: {
                        type: "object",
                        properties: {
                            documentId: {
                                type: "string",
                                description: "ID del documento",
                            },
                            text: {
                                type: "string",
                                description: "Texto a insertar",
                            },
                            index: {
                                type: "number",
                                description: "Posición donde insertar el texto (1 es el inicio del documento)",
                            },
                        },
                        required: ["documentId", "text", "index"],
                    },
                },
                {
                    name: "docs_replace_text",
                    description: "Reemplaza texto en un documento",
                    inputSchema: {
                        type: "object",
                        properties: {
                            documentId: {
                                type: "string",
                                description: "ID del documento",
                            },
                            findText: {
                                type: "string",
                                description: "Texto a buscar",
                            },
                            replaceText: {
                                type: "string",
                                description: "Texto de reemplazo",
                            },
                        },
                        required: ["documentId", "findText", "replaceText"],
                    },
                },
                {
                    name: "docs_format_text",
                    description: "Aplica formato a un rango de texto",
                    inputSchema: {
                        type: "object",
                        properties: {
                            documentId: {
                                type: "string",
                                description: "ID del documento",
                            },
                            startIndex: {
                                type: "number",
                                description: "Índice de inicio del rango",
                            },
                            endIndex: {
                                type: "number",
                                description: "Índice de fin del rango",
                            },
                            bold: {
                                type: "boolean",
                                description: "Aplicar negrita",
                            },
                            italic: {
                                type: "boolean",
                                description: "Aplicar cursiva",
                            },
                            fontSize: {
                                type: "number",
                                description: "Tamaño de fuente en puntos",
                            },
                        },
                        required: ["documentId", "startIndex", "endIndex"],
                    },
                },
                {
                    name: "docs_list_recent",
                    description: "Lista documentos recientes de Google Docs",
                    inputSchema: {
                        type: "object",
                        properties: {
                            maxResults: {
                                type: "number",
                                description: "Máximo de documentos a retornar (default: 10)",
                                default: 10,
                            },
                        },
                    },
                },
                {
                    name: "docs_share",
                    description: "Comparte un documento con otros usuarios",
                    inputSchema: {
                        type: "object",
                        properties: {
                            documentId: {
                                type: "string",
                                description: "ID del documento",
                            },
                            email: {
                                type: "string",
                                description: "Email del usuario con quien compartir",
                            },
                            role: {
                                type: "string",
                                description: "Rol: 'reader', 'writer', o 'commenter'",
                                enum: ["reader", "writer", "commenter"],
                            },
                        },
                        required: ["documentId", "email", "role"],
                    },
                },
                // 🆕 GOOGLE FIT TOOLS
                {
                    name: "fit_get_activity_summary",
                    description: "Obtiene datos agregados de pasos y actividad física en un período. Usa milisegundos de Unix para el tiempo.",
                    inputSchema: {
                        type: "object",
                        properties: {
                            startTimeMillis: {
                                type: "number",
                                description: "Tiempo de inicio del período a consultar (Milisegundos de Unix).",
                            },
                            endTimeMillis: {
                                type: "number",
                                description: "Tiempo de fin del período a consultar (Milisegundos de Unix).",
                            },
                        },
                        required: ["startTimeMillis", "endTimeMillis"],
                    },
                },
                {
                    name: "fit_record_activity_session",
                    description: "Registra manualmente una sesión de actividad física (ej: correr, caminar) en Google Fit.",
                    inputSchema: {
                        type: "object",
                        properties: {
                            activityType: {
                                type: "string",
                                description: "Tipo de actividad (ej: 'running', 'walking', 'yoga'). Usa los códigos de la API de Fit.",
                            },
                            durationMinutes: {
                                type: "number",
                                description: "Duración de la actividad en minutos.",
                            },
                            startTimeMillis: {
                                type: "number",
                                description: "Tiempo de inicio de la actividad (Milisegundos de Unix).",
                            },
                        },
                        required: ["activityType", "durationMinutes", "startTimeMillis"],
                    },
                },
            ],
        }));

        this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
            const { name, arguments: args } = request.params;

            try {
                // Verificar token antes de cada operación
                await this.ensureValidToken();

                switch (name) {
                    // Gmail handlers
                    case "gmail_list_messages":
                        return await this.listGmailMessages(args);
                    case "gmail_get_message":
                        return await this.getGmailMessage(args);
                    case "gmail_send_message":
                        return await this.sendGmailMessage(args);
                    case "gmail_create_draft":
                        return await this.createGmailDraft(args);

                    // Calendar handlers
                    case "calendar_list_events":
                        return await this.listCalendarEvents(args);
                    case "calendar_create_event":
                        return await this.createCalendarEvent(args);
                    case "calendar_update_event":
                        return await this.updateCalendarEvent(args);
                    case "calendar_delete_event":
                        return await this.deleteCalendarEvent(args);
                    case "calendar_find_free_slots":
                        return await this.findFreeSlots(args);

                    // Google Docs handlers
                    case "docs_create":
                        return await this.createDoc(args);
                    case "docs_get":
                        return await this.getDoc(args);
                    case "docs_append_text":
                        return await this.appendText(args);
                    case "docs_insert_text":
                        return await this.insertText(args);
                    case "docs_replace_text":
                        return await this.replaceText(args);
                    case "docs_format_text":
                        return await this.formatText(args);
                    case "docs_list_recent":
                        return await this.listRecentDocs(args);
                    case "docs_share":
                        return await this.shareDoc(args);
                    // 🆕 GOOGLE FIT HANDLERS
                    case "fit_get_activity_summary":
                        return await this.getActivitySummary(args);
                    case "fit_record_activity_session":
                        return await this.recordActivitySession(args);

                    default:
                        throw new Error(`Herramienta desconocida: ${name}`);
                }
            } catch (error) {
                console.error(`Error en ${name}:`, error.message);
                return {
                    content: [
                        {
                            type: "text",
                            text: `Error: ${error.message}`,
                        },
                    ],
                    isError: true,
                };
            }
        });
    }

    // Gmail Methods
    async listGmailMessages(args) {
        const { query = "", maxResults = 10 } = args;
        const res = await this.gmail.users.messages.list({
            userId: "me",
            q: query,
            maxResults,
        });

        const messages = res.data.messages || [];
        const messageDetails = await Promise.all(
            messages.slice(0, 5).map(async (msg) => {
                const detail = await this.gmail.users.messages.get({
                    userId: "me",
                    id: msg.id,
                    format: "metadata",
                    metadataHeaders: ["From", "Subject", "Date"],
                });
                return {
                    id: msg.id,
                    threadId: msg.threadId,
                    headers: detail.data.payload.headers,
                    snippet: detail.data.snippet,
                };
            })
        );

        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify(messageDetails, null, 2),
                },
            ],
        };
    }

    async getGmailMessage(args) {
        const { messageId } = args;
        const res = await this.gmail.users.messages.get({
            userId: "me",
            id: messageId,
            format: "full",
        });

        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify(res.data, null, 2),
                },
            ],
        };
    }

    async sendGmailMessage(args) {
        const { to, subject, body } = args;
        const message = [
            `To: ${to}`,
            `Subject: ${subject}`,
            "",
            body,
        ].join("\n");

        const encodedMessage = Buffer.from(message)
            .toString("base64")
            .replace(/\+/g, "-")
            .replace(/\//g, "_")
            .replace(/=+$/, "");

        const res = await this.gmail.users.messages.send({
            userId: "me",
            requestBody: {
                raw: encodedMessage,
            },
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Email enviado exitosamente. ID del mensaje: ${res.data.id}`,
                },
            ],
        };
    }

    async createGmailDraft(args) {
        const { to, subject, body } = args;
        const message = [
            `To: ${to}`,
            `Subject: ${subject}`,
            "",
            body,
        ].join("\n");

        const encodedMessage = Buffer.from(message)
            .toString("base64")
            .replace(/\+/g, "-")
            .replace(/\//g, "_")
            .replace(/=+$/, "");

        const res = await this.gmail.users.drafts.create({
            userId: "me",
            requestBody: {
                message: {
                    raw: encodedMessage,
                },
            },
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Borrador creado exitosamente. ID del borrador: ${res.data.id}`,
                },
            ],
        };
    }

    // Calendar Methods
    async listCalendarEvents(args) {
        const { maxResults = 10, timeMin, timeMax } = args;
        const res = await this.calendar.events.list({
            calendarId: "primary",
            timeMin: timeMin || new Date().toISOString(),
            timeMax: timeMax,
            maxResults,
            singleEvents: true,
            orderBy: "startTime",
        });

        const events = res.data.items || [];
        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify(events, null, 2),
                },
            ],
        };
    }

    async createCalendarEvent(args) {
        const { summary, description, startDateTime, endDateTime, location, attendees } = args;

        const event = {
            summary,
            description,
            location,
            start: {
                dateTime: startDateTime,
                timeZone: "America/Montevideo",
            },
            end: {
                dateTime: endDateTime,
                timeZone: "America/Montevideo",
            },
        };

        if (attendees && attendees.length > 0) {
            event.attendees = attendees.map((email) => ({ email }));
        }

        const res = await this.calendar.events.insert({
            calendarId: "primary",
            requestBody: event,
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Evento creado exitosamente. ID del evento: ${res.data.id}\nEnlace: ${res.data.htmlLink}`,
                },
            ],
        };
    }

    async updateCalendarEvent(args) {
        const { eventId, summary, description, startDateTime, endDateTime, location } = args;

        // Get existing event first
        const existing = await this.calendar.events.get({
            calendarId: "primary",
            eventId,
        });

        const event = {
            ...existing.data,
            summary: summary || existing.data.summary,
            description: description || existing.data.description,
            location: location || existing.data.location,
        };

        if (startDateTime) {
            event.start = {
                dateTime: startDateTime,
                timeZone: "America/Montevideo",
            };
        }

        if (endDateTime) {
            event.end = {
                dateTime: endDateTime,
                timeZone: "America/Montevideo",
            };
        }

        const res = await this.calendar.events.update({
            calendarId: "primary",
            eventId,
            requestBody: event,
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Evento actualizado exitosamente. ID del evento: ${res.data.id}`,
                },
            ],
        };
    }

    async deleteCalendarEvent(args) {
        const { eventId } = args;
        await this.calendar.events.delete({
            calendarId: "primary",
            eventId,
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Evento eliminado exitosamente. ID del evento: ${eventId}`,
                },
            ],
        };
    }

    async findFreeSlots(args) {
        const { startDate, endDate, duration } = args;

        const res = await this.calendar.events.list({
            calendarId: "primary",
            timeMin: startDate,
            timeMax: endDate,
            singleEvents: true,
            orderBy: "startTime",
        });

        const events = res.data.items || [];
        const freeSlots = [];
        let currentTime = new Date(startDate);
        const endTime = new Date(endDate);

        for (const event of events) {
            const eventStart = new Date(event.start.dateTime || event.start.date);
            const eventEnd = new Date(event.end.dateTime || event.end.date);
            if (eventStart - currentTime >= duration * 60000) {
                freeSlots.push({
                    start: currentTime.toISOString(),
                    end: eventStart.toISOString(),
                    duration: Math.floor((eventStart - currentTime) / 60000),
                });
            }

            currentTime = eventEnd > currentTime ? eventEnd : currentTime;
        }

        if (endTime - currentTime >= duration * 60000) {
            freeSlots.push({
                start: currentTime.toISOString(),
                end: endTime.toISOString(),
                duration: Math.floor((endTime - currentTime) / 60000),
            });
        }

        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify(freeSlots, null, 2),
                },
            ],
        };
    }

    // Google Docs Methods
    async createDoc(args) {
        const { title, content } = args;

        // Create document
        const doc = await this.docs.documents.create({
            requestBody: {
                title,
            },
        });

        const documentId = doc.data.documentId;

        // Add content if provided
        if (content) {
            await this.docs.documents.batchUpdate({
                documentId,
                requestBody: {
                    requests: [
                        {
                            insertText: {
                                location: {
                                    index: 1,
                                },
                                text: content,
                            },
                        },
                    ],
                },
            });
        }

        return {
            content: [
                {
                    type: "text",
                    text: `Documento creado exitosamente.\nID: ${documentId}\nEnlace: https://docs.google.com/document/d/${documentId}/edit`,
                },
            ],
        };
    }

    async getDoc(args) {
        const { documentId } = args;
        const doc = await this.docs.documents.get({
            documentId,
        });

        // Extract text content
        let textContent = "";
        if (doc.data.body && doc.data.body.content) {
            for (const element of doc.data.body.content) {
                if (element.paragraph) {
                    for (const textElement of element.paragraph.elements || []) {
                        if (textElement.textRun) {
                            textContent += textElement.textRun.content;
                        }
                    }
                }
            }
        }

        return {
            content: [
                {
                    type: "text",
                    text: `Título: ${doc.data.title}\n\nContenido:\n${textContent}`,
                },
            ],
        };
    }

    async appendText(args) {
        const { documentId, text } = args;

        // Get document to find the end index
        const doc = await this.docs.documents.get({
            documentId,
        });

        const endIndex = doc.data.body.content[doc.data.body.content.length - 1].endIndex - 1;

        await this.docs.documents.batchUpdate({
            documentId,
            requestBody: {
                requests: [
                    {
                        insertText: {
                            location: {
                                index: endIndex,
                            },
                            text: text,
                        },
                    },
                ],
            },
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Texto añadido exitosamente al documento ${documentId}`,
                },
            ],
        };
    }

    async insertText(args) {
        const { documentId, text, index } = args;

        await this.docs.documents.batchUpdate({
            documentId,
            requestBody: {
                requests: [
                    {
                        insertText: {
                            location: {
                                index,
                            },
                            text,
                        },
                    },
                ],
            },
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Texto insertado exitosamente en la posición ${index}`,
                },
            ],
        };
    }

    async replaceText(args) {
        const { documentId, findText, replaceText } = args;

        await this.docs.documents.batchUpdate({
            documentId,
            requestBody: {
                requests: [
                    {
                        replaceAllText: {
                            containsText: {
                                text: findText,
                                matchCase: true,
                            },
                            replaceText,
                        },
                    },
                ],
            },
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Texto reemplazado exitosamente. "${findText}" → "${replaceText}"`,
                },
            ],
        };
    }

    async formatText(args) {
        const { documentId, startIndex, endIndex, bold, italic, fontSize } = args;

        const textStyle = {};
        if (bold !== undefined) textStyle.bold = bold;
        if (italic !== undefined) textStyle.italic = italic;
        if (fontSize !== undefined) {
            textStyle.fontSize = {
                magnitude: fontSize,
                unit: "PT",
            };
        }

        await this.docs.documents.batchUpdate({
            documentId,
            requestBody: {
                requests: [
                    {
                        updateTextStyle: {
                            range: {
                                startIndex,
                                endIndex,
                            },
                            textStyle,
                            fields: Object.keys(textStyle).join(","),
                        },
                    },
                ],
            },
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Formato aplicado exitosamente al rango ${startIndex}-${endIndex}`,
                },
            ],
        };
    }

    async listRecentDocs(args) {
        const { maxResults = 10 } = args;

        const res = await this.drive.files.list({
            pageSize: maxResults,
            fields: "files(id, name, modifiedTime, webViewLink)",
            q: "mimeType='application/vnd.google-apps.document'",
            orderBy: "modifiedTime desc",
        });

        const files = res.data.files || [];
        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify(files, null, 2),
                },
            ],
        };
    }

    async shareDoc(args) {
        const { documentId, email, role } = args;

        await this.drive.permissions.create({
            fileId: documentId,
            requestBody: {
                type: "user",
                role,
                emailAddress: email,
            },
            sendNotificationEmail: true,
        });

        return {
            content: [
                {
                    type: "text",
                    text: `Documento compartido exitosamente con ${email} (rol: ${role})`,
                },
            ],
        };
    }

    // Google Fit Methods
    async getActivitySummary(args) {
        const { startTimeMillis, endTimeMillis } = args;

        const bucketDuration = 86400000; // 1 día en milisegundos

        const res = await this.fit.users.dataset.aggregate({
            userId: 'me',
            requestBody: {
                aggregateBy: [
                    { dataTypeName: "com.google.step_count.delta" },
                    { dataTypeName: "com.google.calories.expended" }
                ],
                bucketByTime: { durationMillis: bucketDuration },
                startTimeMillis,
                endTimeMillis,
            },
        });

        const summary = res.data.bucket.map(bucket => {
            const stepsData = bucket.dataset.find(d => d.dataSourceId.includes('step_count'));
            const caloriesData = bucket.dataset.find(d => d.dataSourceId.includes('calories.expended'));

            const steps = stepsData && stepsData.point.length ? stepsData.point[0].value[0].intVal : 0;
            const calories = caloriesData && caloriesData.point.length ? caloriesData.point[0].value[0].fpVal : 0;

            return {
                date: new Date(parseInt(bucket.startTimeMillis)).toISOString().split('T')[0],
                steps: steps,
                caloriesExpended: Math.round(calories),
            };
        });

        return {
            content: [{
                type: "text",
                text: JSON.stringify(summary, null, 2),
            }],
        };
    }

    async recordActivitySession(args) {
        const { activityType, durationMinutes, startTimeMillis } = args;

        const durationMillis = durationMinutes * 60 * 1000;
        const endTimeMillis = startTimeMillis + durationMillis;

        // Mapeo de actividades a códigos de Google Fit
        const activityMap = {
            'running': 8,
            'walking': 7,
            'cycling': 1,
            'swimming': 82,
            'yoga': 104,
            'weightlifting': 97,
            'gym': 97,
            'musculacion': 97,
            'sala de musculacion': 97,
        };

        const activityCode = activityMap[activityType.toLowerCase()] || 108; // 108 = other

        // Generar ID único para la sesión
        const sessionId = `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

        const session = {
            id: sessionId,
            name: `${activityType}`,
            description: `Sesión de ${activityType} de ${durationMinutes} minutos`,
            startTimeMillis: startTimeMillis.toString(),
            endTimeMillis: endTimeMillis.toString(),
            activityType: activityCode,
            application: {
                packageName: "com.mcp.googleworkspace",
                version: "1"
            },
        };

        const res = await this.fit.users.sessions.update({
            userId: 'me',
            sessionId: sessionId,
            requestBody: session,
        });

        return {
            content: [{
                type: "text",
                text: `Sesión de actividad registrada exitosamente: ${activityType} de ${durationMinutes} minutos.\nID: ${sessionId}`,
            }],
        };
    }

    async run() {
        await this.initialize();
        const transport = new StdioServerTransport();
        await this.server.connect(transport);
        console.error("✅ Google Workspace MCP Server running on stdio");
    }
}

const server = new GoogleWorkspaceMCPServer();
server.run().catch((error) => {
    console.error("❌ Error fatal durante la inicialización:", error.message);
    process.exit(1);
});