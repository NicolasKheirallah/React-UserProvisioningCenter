import type { JobType } from '../../models';
import type { IWorkflowStepDefinition } from './stepTypes';
import { CLONE_STEPS, ONBOARDING_STEPS } from './steps/onboardingSteps';
import { OFFBOARDING_STEPS } from './steps/offboardingSteps';
import { TRANSFER_STEPS } from './steps/transferSteps';

/** Step pipeline per job type. Bulk rows are submitted as individual Onboard jobs. */
export function stepsForJobType(jobType: JobType): IWorkflowStepDefinition[] {
  switch (jobType) {
    case 'Offboard':
      return OFFBOARDING_STEPS;
    case 'Transfer':
      return TRANSFER_STEPS;
    case 'Clone':
      return CLONE_STEPS;
    default:
      return ONBOARDING_STEPS;
  }
}
