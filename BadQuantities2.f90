program BadQuantities2
implicit none 
! rotating flywheel with torque applied for a duration; find the initial and final kinetic energy
!= KOQRelation :: ROTENERGY = AV ** 2 * MOI
!= KOQRelation :: AV = TORQUE / MOI * TIME
!= KOQ MOI :: I
!= KOQ AV :: ang_vel_init, ang_vel_final
!= KOQ TORQUE :: torque
!= KOQ TIME :: duration
!= KOQ ROTENERGY :: energy_init, energy_final
real, parameter :: I = 6.0, ang_vel_init = 7.0, torque = 8.0, duration = 9.0 
real ::  energy_init, energy_final, ang_vel_final
energy_init = 0.5 * I * ang_vel_init ** 2
! kind-of-quantity correctness, flags this meaningless quantity addition:
energy_final = energy_init + torque
! kind-of-quantity correctness, flags this meaningless quantity multiplication:
energy_final = I / duration ** 2
! correct algorithm for finding final angular velocity, then kinetic energy
ang_vel_final = ang_vel_init + torque / I * duration
energy_final = 0.5 * I * ang_vel_final ** 2
end program BadQuantities2
