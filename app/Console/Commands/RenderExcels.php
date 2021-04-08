<?php

namespace App\Console\Commands;

use Exception;
use Illuminate\Console\Command;
use App\Repositories\EngineFiles;

class RenderExcels extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'render:excels';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Render File excels whether empty or any space';

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        parent::__construct();
        $this->filepaths = [
            base_path("data/") . "Type_A.xls",
            base_path("data/") . "Type_B.xlsx",
        ];

    }

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        foreach($this->filepaths as $filepath) {
            try {
                $engine = new EngineFiles($filepath);
                $tables = $engine->validateAndResult()->toTable();
                $this->table(['Row', 'Error'], $tables);
            } catch (Exception $th) {
                $this->error($th->getMessage());
            }
        }

        $this->info("Job finished!");
        //return 0;
    }
}
