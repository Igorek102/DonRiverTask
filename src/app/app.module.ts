import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { AppComponent } from './app.component';
import {
  ButtonModule, ConfirmDialogModule, ContextMenuModule, DataTableModule, DialogModule, MenuModule, SharedModule,
  TabViewModule
} from 'primeng/primeng';
import {HeaderComponent} from './header.component';
import {ContentComponent} from './content.component';
import {BrowserAnimationsModule, NoopAnimationsModule} from '@angular/platform-browser/animations';
import {FormsModule} from '@angular/forms';


@NgModule({
  declarations: [
    AppComponent, HeaderComponent, ContentComponent
  ],
  imports: [
    BrowserModule,
    ButtonModule,
    TabViewModule,
    DataTableModule,
    SharedModule,
    ConfirmDialogModule,
    BrowserAnimationsModule,
    NoopAnimationsModule,
    MenuModule,
    ContextMenuModule,
    DialogModule,
    FormsModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
