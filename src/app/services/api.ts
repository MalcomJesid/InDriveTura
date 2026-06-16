import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';

export interface ApiResponse {
  data: any[];
}

export interface Perfil {
  id: string;
  nombre_perfil: string;
}

export interface Vehiculo {
  id: string;
  placa: string;
  marca: string;
  modelo: string;
  activo: boolean;
}

export interface Ruta {
  id: string;
  nombre_ruta: string;
  color_hex: string;
}

@Injectable({
  providedIn: 'root'
})
export class Api {

  private baseUrl = 'https://apirecoleccion.gonzaloandreslucio.com/api';

  constructor(private http: HttpClient) {}

  getPerfiles(): Observable<ApiResponse> {
    return this.http.get<ApiResponse>(`${this.baseUrl}/perfiles/todas`);
  }

 getVehiculos(perfilId: string): Observable<ApiResponse> {
  return this.http.get<ApiResponse>(`${this.baseUrl}/vehiculos?perfil_id=${perfilId}`);
}

  getRutas(): Observable<ApiResponse> {
    return this.http.get<ApiResponse>(`${this.baseUrl}/rutas/todas`);
  }

}