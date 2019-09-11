import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import {BrowserAnimationsModule} from '@angular/platform-browser/animations';
import { AppComponent } from './app.component';
import { DataTableComponent } from './components/data-table/data-table.component';
import {
  MatButtonModule,
  MatIconModule,
  MatInputModule,
  MatNativeDateModule,
  MatTableModule
} from '@angular/material';
import { ExporterService } from './services/exporter.service';

@NgModule({
  declarations: [
    AppComponent,
    DataTableComponent
  ],
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    MatButtonModule,
    MatIconModule,
    MatInputModule,
    MatNativeDateModule,
    MatTableModule
  ],
  providers: [ExporterService],
  bootstrap: [AppComponent]
})
export class AppModule { }
