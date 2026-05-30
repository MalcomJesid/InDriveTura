import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';

export interface ApiResponse {
  data: Perfil[];
}
export interface Perfil {
  id: string;
  nombre_perfil: string;
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

}