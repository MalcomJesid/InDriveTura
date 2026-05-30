import { Component, OnInit } from '@angular/core';
import { Api, Perfil } from '../services/api';

@Component({
  selector: 'app-home',
  templateUrl: 'home.page.html',
  styleUrls: ['home.page.scss'],
  standalone: false,
})
export class HomePage implements OnInit {

  perfiles: Perfil[] = [];
  cargando: boolean = true;
  error: string = '';

  constructor(private apiService: Api) {}

  ngOnInit() {
    this.cargarPerfiles();
  }

  cargarPerfiles() {
    this.apiService.getPerfiles().subscribe({
      next: (respuesta) => {
        console.log('Respuesta de la API:', respuesta);
        this.perfiles = respuesta.data;
        this.cargando = false;
    },
      error: (err) => {
        this.error = 'No se pudo conectar con la API';
        this.cargando = false;
        console.error(err);
      }

    });
  }

}