import {Component, ViewEncapsulation} from '@angular/core';
import {ConfirmationService} from 'primeng/primeng';

@Component({
  selector: 'app-header',
  templateUrl: 'header.component.html',
  styleUrls: ['header.component.css'],
  encapsulation: ViewEncapsulation.None,
  providers: [ConfirmationService]
})
export class HeaderComponent {
  displayInfo = false;
  btnNames: string[] = [
    'NBN ISP Report', 'DSL ISP Report', 'Configuration'
  ]
  constructor(private confirmationService: ConfirmationService) {}
  confirm(index: number) {
    this.confirmationService.confirm({
      message: 'Item \'' + this.btnNames[index] + '\' has been activated',
      header: 'Confirmation Dialog',
      icon: 'fa-info',
      rejectVisible: false,
      key: 'conf'
    });
  }
  onLogout(): void {
    this.confirmationService.confirm({
      message: 'Are you sure you want to log out?',
      header: 'Log Out',
      icon: 'fa-question',
      rejectVisible: true,
      key: 'logout',
      accept: () => {
        this.displayInfo = true;
      }
    });
  }
  closeDialog() {
    this.displayInfo = false;
  }
}
