// Main JavaScript for Exam Seating Generator

document.addEventListener('DOMContentLoaded', function() {
    // Initialize sidebar functionality
    const sidebarToggle = document.getElementById('sidebarToggle');
    const sidebar = document.getElementById('sidebar');
    
    if (sidebarToggle && sidebar) {
        sidebarToggle.addEventListener('click', function() {
            sidebar.classList.toggle('show');
            
            // Create overlay for mobile
            let overlay = document.querySelector('.sidebar-overlay');
            if (!overlay) {
                overlay = document.createElement('div');
                overlay.className = 'sidebar-overlay';
                document.body.appendChild(overlay);
                
                overlay.addEventListener('click', function() {
                    sidebar.classList.remove('show');
                    overlay.classList.remove('show');
                });
            }
            
            if (sidebar.classList.contains('show')) {
                overlay.classList.add('show');
            } else {
                overlay.classList.remove('show');
            }
        });
    }
    
    // Initialize tooltips
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    const tooltipList = tooltipTriggerList.map(function(tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });
    
    // Auto-dismiss alerts after 5 seconds
    const alerts = document.querySelectorAll('.alert-dismissible');
    alerts.forEach(function(alert) {
        setTimeout(function() {
            const bsAlert = new bootstrap.Alert(alert);
            if (bsAlert) {
                bsAlert.close();
            }
        }, 5000);
    });
    
    // Form validation improvements
    const forms = document.querySelectorAll('form');
    forms.forEach(function(form) {
        form.addEventListener('submit', function(event) {
            if (!form.checkValidity()) {
                event.preventDefault();
                event.stopPropagation();
            } else {
                // Show loading state for generate button
                const submitBtn = form.querySelector('button[type="submit"]');
                if (submitBtn && submitBtn.id === 'generateBtn') {
                    submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2" role="status"></span>Generating...';
                    submitBtn.disabled = true;
                }
            }
            form.classList.add('was-validated');
        });
    });
    
    // File input validation
    const fileInputs = document.querySelectorAll('input[type="file"]');
    fileInputs.forEach(function(input) {
        input.addEventListener('change', function(event) {
            const file = event.target.files[0];
            if (file) {
                const fileName = file.name;
                const fileSize = (file.size / 1024 / 1024).toFixed(2); // MB
                const allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                                    'application/vnd.ms-excel'];
                
                if (!allowedTypes.includes(file.type)) {
                    showAlert('Please select a valid Excel file (.xlsx or .xls)', 'error');
                    input.value = '';
                    return;
                }
                
                if (file.size > 16 * 1024 * 1024) { // 16MB limit
                    showAlert('File size must be less than 16MB', 'error');
                    input.value = '';
                    return;
                }
                
                // Update file input label if exists
                const label = input.parentElement.querySelector('label');
                if (label) {
                    label.textContent = `${fileName} (${fileSize}MB)`;
                }
            }
        });
    });
    
    // Number input validation
    const numberInputs = document.querySelectorAll('input[type="number"]');
    numberInputs.forEach(function(input) {
        input.addEventListener('input', function() {
            const value = parseInt(this.value);
            const min = parseInt(this.getAttribute('min') || '0');
            const max = parseInt(this.getAttribute('max') || '999');
            
            if (value < min) {
                this.value = min;
            } else if (value > max) {
                this.value = max;
            }
            
            // Update capacity calculation if this is room configuration
            if (this.id === 'rows' || this.id === 'columns' || this.id === 'extra_desks') {
                updateRoomCapacity();
            }
        });
    });
    
    // Room capacity calculation
    function updateRoomCapacity() {
        const rowsInput = document.getElementById('rows');
        const columnsInput = document.getElementById('columns');
        const extraDesksInput = document.getElementById('extra_desks');
        
        if (rowsInput && columnsInput && extraDesksInput) {
            const rows = parseInt(rowsInput.value) || 0;
            const columns = parseInt(columnsInput.value) || 0;
            const extraDesks = parseInt(extraDesksInput.value) || 0;
            
            const gridCapacity = rows * columns * 2;
            const extraCapacity = extraDesks * 2;
            const totalCapacity = gridCapacity + extraCapacity;
            
            // Update capacity displays
            const gridCapacityEl = document.getElementById('grid-capacity');
            const extraCapacityEl = document.getElementById('extra-capacity');
            const totalCapacityEl = document.getElementById('total-capacity');
            
            if (gridCapacityEl) gridCapacityEl.textContent = gridCapacity;
            if (extraCapacityEl) extraCapacityEl.textContent = extraCapacity;
            if (totalCapacityEl) totalCapacityEl.textContent = totalCapacity;
        }
    }
    
    // Class selection helpers
    const classCheckboxContainer = document.getElementById('class-checkboxes');
    if (classCheckboxContainer) {
        // Add select all/none functionality
        const selectAllBtn = document.createElement('button');
        selectAllBtn.type = 'button';
        selectAllBtn.className = 'btn btn-sm btn-outline-secondary me-2 mb-2';
        selectAllBtn.innerHTML = '<i class="bi bi-check-all"></i> All';
        selectAllBtn.onclick = function() {
            const checkboxes = classCheckboxContainer.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(cb => cb.checked = true);
        };
        
        const selectNoneBtn = document.createElement('button');
        selectNoneBtn.type = 'button';
        selectNoneBtn.className = 'btn btn-sm btn-outline-secondary mb-2';
        selectNoneBtn.innerHTML = '<i class="bi bi-x-square"></i> None';
        selectNoneBtn.onclick = function() {
            const checkboxes = classCheckboxContainer.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(cb => cb.checked = false);
        };
        
        // Insert buttons at the beginning
        const firstChild = classCheckboxContainer.firstChild;
        if (firstChild && !firstChild.textContent.includes('No classes')) {
            classCheckboxContainer.insertBefore(selectAllBtn, firstChild);
            classCheckboxContainer.insertBefore(selectNoneBtn, firstChild);
        }
    }
    
    // Seating grid interactions
    const deskCards = document.querySelectorAll('.desk-card');
    deskCards.forEach(function(card) {
        card.addEventListener('click', function() {
            // Toggle detailed view or show modal with desk information
            const desk = this;
            const deskId = desk.querySelector('.desk-header').textContent;
            const seats = desk.querySelectorAll('.seat');
            
            // Create modal content
            let modalContent = `<h6>Desk ${deskId} Details</h6>`;
            
            seats.forEach(function(seat, index) {
                const seatLetter = index === 0 ? 'A' : 'B';
                const isOccupied = seat.classList.contains('bg-success');
                
                if (isOccupied) {
                    const studentInfo = seat.innerHTML;
                    modalContent += `<p><strong>Seat ${seatLetter}:</strong><br>${studentInfo}</p>`;
                } else {
                    modalContent += `<p><strong>Seat ${seatLetter}:</strong> <span class="text-muted">Empty</span></p>`;
                }
            });
            
            // Show details in a simple alert for now (can be enhanced with modal)
            showAlert(modalContent, 'info');
        });
    });
});

// Utility function to show custom alerts
function showAlert(message, type = 'info') {
    const alertContainer = document.querySelector('.container .alert');
    const parentContainer = alertContainer ? alertContainer.parentElement : document.querySelector('.container');
    
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type === 'error' ? 'danger' : type} alert-dismissible fade show`;
    alertDiv.innerHTML = `
        <i class="bi bi-${type === 'error' ? 'exclamation-triangle' : type === 'success' ? 'check-circle' : 'info-circle'}"></i>
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;
    
    if (parentContainer) {
        parentContainer.insertBefore(alertDiv, parentContainer.firstChild);
        
        // Auto-dismiss after 5 seconds
        setTimeout(function() {
            const bsAlert = new bootstrap.Alert(alertDiv);
            if (bsAlert) {
                bsAlert.close();
            }
        }, 5000);
    }
}

// Utility function to format numbers with thousands separator
function formatNumber(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// Export functionality helpers
function downloadFile(url, filename) {
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Print functionality
function printSeatingPlan() {
    window.print();
}

// Add print button to preview page if it exists
document.addEventListener('DOMContentLoaded', function() {
    const previewPage = document.querySelector('h1');
    if (previewPage && previewPage.textContent.includes("Seating Plan Preview")) {
        const printBtn = document.createElement('button');
        printBtn.className = 'btn btn-outline-secondary';
        printBtn.innerHTML = '<i class="bi bi-printer"></i> Print';
        printBtn.onclick = printSeatingPlan;
        
        const buttonContainer = previewPage.nextElementSibling;
        if (buttonContainer && buttonContainer.querySelector('.btn')) {
            buttonContainer.appendChild(printBtn);
        }
    }
});
