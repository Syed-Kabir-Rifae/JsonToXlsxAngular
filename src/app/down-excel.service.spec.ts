import { TestBed } from '@angular/core/testing';

import { DownExcelService } from './down-excel.service';

describe('DownExcelService', () => {
  let service: DownExcelService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(DownExcelService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
