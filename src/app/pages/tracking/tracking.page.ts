import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { StorageService } from '../../services/storage';

@Component({
  selector: 'app-tracking',
  templateUrl: './tracking.page.html',
  styleUrls: ['./tracking.page.scss'],
  standalone: false,
})
export class TrackingPage implements OnInit {

  perfilNombre: string = '';
  vehiculoPlaca: string = '';
  rutaNombre: string = '';

  constructor(
    private storageService: StorageService,
    private router: Router
  ) {}

  async ngOnInit() {
    this.perfilNombre  = await this.storageService.obtener('perfil_nombre');
    this.vehiculoPlaca = await this.storageService.obtener('vehiculo_placa');
    this.rutaNombre    = await this.storageService.obtener('ruta_nombre');
  }

  volverAtras() {
    this.router.navigate(['/routes']);
  }

}