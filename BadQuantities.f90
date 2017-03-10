program BadQuantities1
implicit none 
! rotating flywheel with torque applied for a duration
! find the initial and final kinetic energy
! annotated with Camfort units declarations
!= unit kg*m**2 :: I
!= unit kg m**2/s**2 :: torque
!= unit s :: duration
!= unit 1/s :: ang_vel_init, ang_vel_final
!= unit kg m**2/s**2 :: energy_init, energy_final
real, parameter :: I = 6.0, ang_vel_init = 7.0, torque = 8.0, duration = 9.0 
real ::  energy_init, energy_final, ang_vel_final
energy_init = 0.5 * I * ang_vel_init ** 2
! unit correctness, allows this meaningless quantity addition:
energy_final = energy_init + torque
! unit correctness, allows this meaningless quantity multiplication:
energy_final = I / duration ** 2
! correct algorithm for finding final angular velocity, then kinetic energy
ang_vel_final = ang_vel_init + torque / I * duration
energy_final = 0.5 * I * ang_vel_final ** 2
end program BadQuantities1
