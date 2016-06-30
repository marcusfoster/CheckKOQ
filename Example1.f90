program example1
implicit none 
! This is an example target source code for unit and quantity checking.
! This has unit annotations for checking by 'Camfort', a Fortran source-code units-of-measure checker described in:
! M. Contrastin, et al., "Units-of-Measure Correctness in Fortran Programs," Computing in Science & Engineering, vol. 18, pp. 102-107, 2016.
! The source code must be annotated with comments as follows:
!= unit <unit_name> [:: variable_name]  

! rotating flywheel with torque applied for a duration
! find the initial and final kinetic energy
!= unit kg*m**2 :: I
!= unit 1/s :: ang_vel_init
!= unit kg*m**2/s**2 :: torque
!= unit s :: duration
real, parameter :: I = 6.0, ang_vel_init = 7.0, torque = 8.0, duration = 9.0 
!= unit kg*m**2/s**2 :: energy_init, energy_final
real ::  energy_init, energy_final
energy_init = 0.5 * I * ang_vel_init ** 2
! unit correctness, allows this meaningless quantity addition:
energy_final = energy_init + torque
! unit correctness, allows this meaningless quantity multiplication:
energy_final = I / duration ** 2
end program example1
