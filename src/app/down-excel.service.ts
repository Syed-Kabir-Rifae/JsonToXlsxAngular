import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root'
})

export class DownExcelService {
  baseUrl: String = "http://localhost:8080/export/excel";

  constructor(private http: HttpClient) { }

exportExcel(): Observable<Blob>{
  return this.http.get(`${this.baseUrl}/export/excel`,{responseType:'blob'});
} 
}
