import { Injectable } from '@angular/core';
import { Storage } from '@ionic/storage-angular';

@Injectable({
  providedIn: 'root'
})
export class StorageService {

  constructor(private storage: Storage) {}

  async inicializar() {
    await this.storage.create();
  }

  async guardar(clave: string, valor: any) {
    await this.storage.set(clave, valor);
  }

  async obtener(clave: string): Promise<any> {
    return await this.storage.get(clave);
  }

  async eliminar(clave: string) {
    await this.storage.remove(clave);
  }

  async limpiarTodo() {
    await this.storage.clear();
  }

}