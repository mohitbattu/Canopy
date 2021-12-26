pipeline{
  agent any
  stages{
    stage('build'){
      steps{
        echo "building the application.."
        nodejs('Node-10.7'){
          sh 'yarn install'
          
        }
        echo "Application built"
      }
    }
    stage('test'){
      steps{
        echo "testing the application.."
        echo "Tested the Application"
        withGradle(){
          sh './gradlew -v'
        }
      }
    }
    stage('deploy'){
      steps{
        echo "deploying the application.."
        echo "Deployed the Application"
      }
    }
  }
}
