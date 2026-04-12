const ARM_BASE = 'https://management.azure.com';

export class ArmClient {
  constructor(private readonly getToken: () => Promise<string>) {}

  async get<T>(path: string): Promise<T> {
    const token = await this.getToken();
    const response = await fetch(`${ARM_BASE}${path}`, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'Accept-Language': 'en-US'
      }
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`ARM GET ${path} failed: ${response.status} ${body}`);
    }

    return response.json() as Promise<T>;
  }

  async post<T>(path: string, payload: object): Promise<T> {
    const token = await this.getToken();
    const response = await fetch(`${ARM_BASE}${path}`, {
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
      throw new Error(`ARM POST ${path} failed: ${response.status} ${body}`);
    }

    return response.json() as Promise<T>;
  }

  async put<T>(path: string, payload: object): Promise<T> {
    const token = await this.getToken();
    const response = await fetch(`${ARM_BASE}${path}`, {
      method: 'PUT',
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
      throw new Error(`ARM PUT ${path} failed: ${response.status} ${body}`);
    }

    return response.json() as Promise<T>;
  }
}
