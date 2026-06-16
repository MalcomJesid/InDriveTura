import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { Api, Perfil } from '../services/api';
import { StorageService } from '../services/storage';

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

  constructor(
    private apiService: Api,
    private storageService: StorageService,
    private router: Router
  ) {}

  ngOnInit() {
    this.cargarPerfiles();
  }

  cargarPerfiles() {
    this.apiService.getPerfiles().subscribe({
      next: (respuesta) => {
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

  async seleccionarPerfil(perfil: Perfil) {
    await this.storageService.guardar('perfil_id', perfil.id);
    await this.storageService.guardar('perfil_nombre', perfil.nombre_perfil);
    this.router.navigate(['/vehicles']);
  }

}