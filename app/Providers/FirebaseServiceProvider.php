<?php

namespace App\Providers;

use Illuminate\Support\ServiceProvider;
use Kreait\Firebase\Factory;

class FirebaseServiceProvider extends ServiceProvider
{
    /**
     * Register services.
     *
     * @return void
     */
    public function register()
    {
        $this->app->singleton('firebase.firestore', function ($app) {
            $factory = (new Factory)->withServiceAccount(storage_path('app/firebase/app-webding-firebase.json'));
            $firestore = $factory->createFirestore();
            return $firestore->database();
        });
    }

    /**
     * Bootstrap services.
     *
     * @return void
     */
    public function boot()
    {
        //
    }
}
