const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export class GraphClient {
  constructor(private readonly getToken: () => Promise<string>) {}

  async get<T>(path: string): Promise<T> {
    const token = await this.getToken();
    const response = await fetch(`${GRAPH_BASE}${path}`, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'Accept-Language': 'en-US'
      }
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Graph GET ${path} failed: ${response.status} ${body}`);
    }

    return response.json() as Promise<T>;
  }

  async post<T>(path: string, payload: object): Promise<T> {
    const token = await this.getToken();
    const response = await fetch(`${GRAPH_BASE}${path}`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
        'Accept-Language': 'en-US'
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Graph POST ${path} failed: ${response.status} ${body}`);
    }

    return response.json() as Promise<T>;
  }
}
