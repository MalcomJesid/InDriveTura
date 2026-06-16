import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { Api, Ruta } from '../../services/api';
import { StorageService } from '../../services/storage';

@Component({
  selector: 'app-routes',
  templateUrl: './routes.page.html',
  styleUrls: ['./routes.page.scss'],
  standalone: false,
})
export class RoutesPage implements OnInit {

  rutas: Ruta[] = [];
  cargando: boolean = true;
  error: string = '';
  perfilNombre: string = '';
  vehiculoPlaca: string = '';

  constructor(
    private apiService: Api,
    private storageService: StorageService,
    private router: Router
  ) {}

  async ngOnInit() {
    this.perfilNombre = await this.storageService.obtener('perfil_nombre');
    this.vehiculoPlaca = await this.storageService.obtener('vehiculo_placa');
    this.cargarRutas();
  }

  cargarRutas() {
    this.apiService.getRutas().subscribe({
      next: (respuesta) => {
        this.rutas = respuesta.data;
        this.cargando = false;
      },
      error: (err) => {
        this.error = 'No se pudo cargar las rutas';
        this.cargando = false;
        console.error(err);
      }
    });
  }

  async seleccionarRuta(ruta: Ruta) {
    await this.storageService.guardar('ruta_id', ruta.id);
    await this.storageService.guardar('ruta_nombre', ruta.nombre_ruta);
    this.router.navigate(['/tracking']);
  }

  volverAtras() {
    this.router.navigate(['/vehicles']);
  }

}