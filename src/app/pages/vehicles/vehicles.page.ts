import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { Api, Vehiculo } from '../../services/api';
import { StorageService } from '../../services/storage';

@Component({
  selector: 'app-vehicles',
  templateUrl: './vehicles.page.html',
  styleUrls: ['./vehicles.page.scss'],
  standalone: false,
})
export class VehiclesPage implements OnInit {

  vehiculos: Vehiculo[] = [];
  cargando: boolean = true;
  error: string = '';
  perfilNombre: string = '';

  constructor(
    private apiService: Api,
    private storageService: StorageService,
    private router: Router
  ) {}

  async ngOnInit() {
  this.perfilNombre = await this.storageService.obtener('perfil_nombre');
  const perfilId = await this.storageService.obtener('perfil_id');
  this.cargarVehiculos(perfilId);
}

cargarVehiculos(perfilId: string) {
  this.apiService.getVehiculos(perfilId).subscribe({
    next: (respuesta) => {
      this.vehiculos = respuesta.data;
      this.cargando = false;
    },
    error: (err) => {
      this.error = 'No se pudo cargar los vehículos';
      this.cargando = false;
      console.error(err);
    }
  });
}

  async seleccionarVehiculo(vehiculo: Vehiculo) {
    await this.storageService.guardar('vehiculo_id', vehiculo.id);
    await this.storageService.guardar('vehiculo_placa', vehiculo.placa);
    this.router.navigate(['/routes']);
  }

  volverAtras() {
    this.router.navigate(['/home']);
  }

}